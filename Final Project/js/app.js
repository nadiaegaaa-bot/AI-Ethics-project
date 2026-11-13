/* Bulk Energy Intelligence dashboard. Every number is computed from window.BULK_DATA (js/data.js), generated from the Excel workbook. */
(() => {
  const ICONS = {"lock": "<rect x=\"5\" y=\"11\" width=\"14\" height=\"10\" rx=\"2\"/><path d=\"M8 11V7a4 4 0 0 1 8 0v4\"/>", "home": "<path d=\"M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z\"/>", "sites": "<rect x=\"3\" y=\"7\" width=\"7\" height=\"14\"/><rect x=\"10\" y=\"3\" width=\"11\" height=\"18\"/><path d=\"M13 7h2M17 7h1M13 11h2M17 11h1M13 15h2M17 15h1\"/>", "bolt": "<path d=\"M13 2L4 14h7l-1 8 9-12h-7z\"/>", "spark": "<path d=\"M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8zM19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z\"/>", "db": "<ellipse cx=\"12\" cy=\"5\" rx=\"8\" ry=\"3\"/><path d=\"M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3\"/>", "chart": "<path d=\"M3 3v18h18M8 17V11M13 17V7M18 17v-4\"/>", "file": "<path d=\"M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z\"/><path d=\"M14 3v5h5M9 13h6M9 17h6\"/>", "gear": "<circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z\"/>", "warn": "<path d=\"M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0zM12 9v4M12 17h.01\"/>", "bulb": "<path d=\"M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.1V17h6v-.2c0-.8.4-1.6 1-2.1A7 7 0 0 0 12 2z\"/>", "check": "<path d=\"M20 6L9 17l-5-5\"/>", "info": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M12 16v-4M12 8h.01\"/>", "target": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><circle cx=\"12\" cy=\"12\" r=\"5\"/><circle cx=\"12\" cy=\"12\" r=\"1\"/>", "leaf": "<path d=\"M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.5 19 2c1 2 2 4.2 2 8 0 5.5-4.8 10-10 10zM2 21c0-3 1.9-5.4 5.1-6\"/>", "link": "<path d=\"M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7\"/>", "down": "<path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3\"/>", "up": "<path d=\"M3 17l6-6 4 4 8-8M15 7h6v6\"/>", "dn": "<path d=\"M3 7l6 6 4-4 8 8M15 17h6v-6\"/>", "chev": "<path d=\"M9 18l6-6-6-6\"/>", "arrow": "<path d=\"M5 12h14M13 6l6 6-6 6\"/>", "db2": "<ellipse cx=\"12\" cy=\"5\" rx=\"8\" ry=\"3\"/><path d=\"M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 11v8c0 1.7 3.6 3 8 3s8-1.3 8-3v-8\"/>", "book": "<path d=\"M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z\"/><path d=\"M4 19.5A2.5 2.5 0 0 0 6.5 22H20v-5\"/>", "ext": "<path d=\"M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14L21 3\"/>"};
  const ic = (n, s = 15) => `<svg class="icon" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICONS[n] || ''}</svg>`;

  const D = window.BULK_DATA;
  const C = { purple: '#7a5cf5', blue: '#4fa3f0', other: '#d9dbe3', green: '#3ecf72', amber: '#f6b93b', red: '#ff5c5c', muted: '#9aa0b8', grid: '#23273b' };
  const SITE_INFO = {
    'N01': { location: 'Vennesla (Kristiansand)', country: 'Norway' },
    'OS-IX': { location: 'Oslo', country: 'Norway' },
    'DK01': { location: 'Esbjerg', country: 'Denmark' },
  };
  const KNOWN = ['OS-IX', 'N01', 'DK01'];
  const SITES = [...KNOWN.filter(s => D.monthly.some(r => r.site === s)), ...[...new Set(D.monthly.map(r => r.site))].filter(s => !KNOWN.includes(s)).sort()];
  SITES.forEach(s => { if (!SITE_INFO[s]) SITE_INFO[s] = { location: s, country: '' }; });
  const UPLOADS = D.uploads || [];
  const MONTHS = [...new Set(D.monthly.map(r => r.month))].sort();
  const LAST = MONTHS.at(-1);
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mShort = m => `${MON[+m.slice(5) - 1]} ${m.slice(2, 4)}`;
  const mLong = m => `${MON[+m.slice(5) - 1]} ${m.slice(0, 4)}`;
  const fact = (entity, metric) => D.facts.find(f => f.entity === entity && f.metric === metric);
  const N01_CAPACITY = fact('N01 Vennesla', 'Operating IT capacity');
  const N01_PLANNED = fact('N01 Vennesla', 'Total planned IT capacity');
  const PUE_TARGET = (fact('Bulk Data Centers', 'New build design PUE objective') || { value: 1.2 }).value;
  const THRESH = D.reviewThreshold;
  const zone = s => (D.monthly.find(r => r.site === s) || {}).zone || '';

  // Combine competitors reported as separate sites without a portfolio figure (e.g. Verne Iceland + Verne Finland) into one entry.
  (function mergeMultiSitePeers() {
    const groups = {};
    D.benchmarks.forEach(b => { if (!/^Bulk/i.test(b.operator) && !/TARGET/i.test(b.status || '')) (groups[b.operator] ||= []).push(b); });
    Object.entries(groups).forEach(([op, list]) => {
      const nums = list.filter(b => typeof b.pue === 'number');
      if (list.length < 2 || nums.length < 2 || list.some(b => /portfolio/i.test(b.facility || ''))) return;
      const pues = nums.map(b => b.pue), years = [...new Set(nums.map(b => b.year))];
      const parts = nums.map(b => ({ facility: b.facility, pue: b.pue, year: b.year }));
      const merged = {
        operator: op, facility: nums.map(b => b.facility.replace(/ campus$/i, '')).join(' + '), year: years.join('/'),
        pue: +(pues.reduce((t, v) => t + v, 0) / pues.length).toFixed(3), parts, merged: true,
        pueType: `Simple average of ${parts.map(p => `${p.facility} ${p.pue.toFixed(2)}`).join(' and ')}. Not energy-weighted, because energy per site is not disclosed.`,
        energyMWh: nums.every(b => typeof b.energyMWh === 'number') ? nums.reduce((t, b) => t + b.energyMWh, 0) : null,
        energyScope: nums[0].energyScope, status: nums.every(b => b.status === 'VERIFIED') ? 'VERIFIED (combined)' : nums[0].status,
        relevance: nums[0].relevance, source: nums[0].source,
        cautions: nums.map(b => `${b.facility}: ${b.cautions}`).join(' ') + ' Combined figure calculated in the dashboard.',
      };
      const at = D.benchmarks.indexOf(list[0]);
      D.benchmarks = D.benchmarks.filter(b => !list.includes(b));
      D.benchmarks.splice(at, 0, merged);
      const names = [...new Set(D.peerMonthly.map(r => r.peer))].filter(n => n.startsWith(op + ' '));
      if (names.length > 1) {
        const byM = {};
        D.peerMonthly.filter(r => names.includes(r.peer)).forEach(r => {
          const m = byM[r.month] ||= { month: r.month, peer: op, refYear: r.refYear, refPUE: merged.pue, totalMWh: 0, itMWh: 0, provenance: r.provenance };
          m.totalMWh += r.totalMWh; m.itMWh += r.itMWh;
        });
        const first = D.peerMonthly.findIndex(r => names.includes(r.peer));
        D.peerMonthly = D.peerMonthly.filter(r => !names.includes(r.peer));
        D.peerMonthly.splice(first, 0, ...Object.values(byM).sort((x, y) => x.month.localeCompare(y.month)).map(m => ({ ...m, pue: +(m.totalMWh / m.itMWh).toFixed(4) })));
      }
    });
  })();

  const app = document.getElementById('app');
  const store = {
    get: (k, d) => { try { return JSON.parse(localStorage.getItem('bulk2.' + k)) ?? d; } catch { return d; } },
    set: (k, v) => localStorage.setItem('bulk2.' + k, JSON.stringify(v)),
  };
  const state = { period: store.get('period', 'ytd'), site: 'N01', siteTab: 'performance', benchTab: 'reported', dataTab: 'facts' };
  let charts = [];

  Chart.defaults.color = C.muted;
  Chart.defaults.font.family = 'Manrope, system-ui, sans-serif';
  Chart.defaults.font.size = 11;
  Chart.defaults.borderColor = C.grid;
  Chart.defaults.plugins.legend.labels.boxWidth = 10;

  // Overlay plugin: dashed horizontal line(s) and shaded x-range boxes.
  Chart.register({
    id: 'overlays',
    beforeDatasetsDraw(chart, _a, o) {
      const { ctx, chartArea: ca, scales: { x, y } } = chart;
      if (!x || !y || !o) return;
      const step = x.getPixelForValue(1) - x.getPixelForValue(0);
      (o.boxes || []).forEach(b => {
        const x0 = x.getPixelForValue(b.from) - step / 2, x1 = x.getPixelForValue(b.to) + step / 2;
        ctx.save(); ctx.fillStyle = b.color || 'rgba(122,92,245,0.2)'; ctx.fillRect(x0, ca.top, x1 - x0, ca.bottom - ca.top);
        if (b.label) { ctx.fillStyle = '#fff'; ctx.font = '11px Manrope, sans-serif'; ctx.textAlign = 'center'; ctx.fillText(b.label, (x0 + x1) / 2, ca.top + 14); }
        ctx.restore();
      });
    },
    afterDatasetsDraw(chart, _a, o) {
      const { ctx, chartArea: ca, scales: { y } } = chart;
      if (!y || !o) return;
      (o.hlines || []).forEach(h => {
        const py = y.getPixelForValue(h.y);
        if (py < ca.top || py > ca.bottom) return;
        ctx.save(); ctx.strokeStyle = h.color || '#7c819c'; ctx.setLineDash([5, 5]); ctx.beginPath(); ctx.moveTo(ca.left, py); ctx.lineTo(ca.right, py); ctx.stroke(); ctx.setLineDash([]);
        if (h.label) {
          ctx.font = '10px Manrope, sans-serif'; ctx.textAlign = 'right';
          const w = ctx.measureText(h.label).width + 8, rx = ca.right - 4;
          ctx.fillStyle = 'rgba(20,23,38,0.85)'; ctx.fillRect(rx - w + 4, py - 15, w, 13); ctx.fillStyle = '#c9ccda'; ctx.fillText(h.label, rx, py - 5);
        }
        ctx.restore();
      });
    },
  });

  /* ---------- periods ---------- */
  const PERIODS = {
    last: 'Latest month',
    '3m': 'Last 3 months',
    ytd: `Year to date ${LAST.slice(0, 4)}`,
    '12m': 'Last 12 months',
    y2025: 'Full year 2025',
    all: `All data (${mShort(MONTHS[0])} – ${mShort(LAST)})`,
  };
  function periodMonths(p = state.period) {
    const n = MONTHS.length, i = MONTHS.length;
    const lastN = k => ({ cur: MONTHS.slice(i - k), prev: i - 2 * k >= 0 ? MONTHS.slice(i - 2 * k, i - k) : null, prevLabel: 'vs. previous' });
    if (p === 'last') return lastN(1);
    if (p === '3m') return lastN(3);
    if (p === '12m') return lastN(12);
    if (p === 'ytd') {
      const y = LAST.slice(0, 4), cur = MONTHS.filter(m => m.startsWith(y));
      const prev = cur.map(m => `${+y - 1}${m.slice(4)}`);
      return { cur, prev: prev.every(m => MONTHS.includes(m)) ? prev : null, prevLabel: `vs. same period ${+y - 1}` };
    }
    if (p === 'y2025') return { cur: MONTHS.filter(m => m.startsWith('2025')), prev: null };
    return { cur: MONTHS.slice(), prev: null };
  }

  /* ---------- data helpers ---------- */
  const rows = (months, sites = SITES) => D.monthly.filter(r => months.includes(r.month) && sites.includes(r.site));
  const agg = rs => {
    const a = rs.reduce((t, r) => ({ it: t.it + r.itMWh, cool: t.cool + r.coolingMWh, other: t.other + r.otherMWh, total: t.total + r.totalMWh, cost: t.cost + r.costNOK, exp: t.exp + r.expectedCoolingMWh, hours: t.hours + r.hours }), { it: 0, cool: 0, other: 0, total: 0, cost: 0, exp: 0, hours: 0 });
    a.pue = a.it ? a.total / a.it : 0;
    a.price = a.total ? a.cost / a.total : 0; // NOK/MWh
    a.coolShare = a.total ? a.cool / a.total : 0;
    a.residual = a.exp ? (a.cool - a.exp) / a.exp : 0;
    return a;
  };
  const cur = sites => agg(rows(periodMonths().cur, sites));
  const prev = sites => { const p = periodMonths().prev; return p ? agg(rows(p, sites)) : null; };
  const flagged = (months = MONTHS, sites = SITES) => rows(months, sites).filter(r => r.alert === 'REVIEW').sort((a, b) => b.month.localeCompare(a.month));
  const excessNOK = r => Math.max(r.residualMWh, 0) * r.priceNOKMWh;

  /* ---------- format helpers ---------- */
  const gwh = mwh => `${(mwh / 1000).toFixed(1)} GWh`;
  const nokM = n => `NOK ${(n / 1e6).toFixed(1)}M`;
  const pctS = (x, d = 1) => `${x >= 0 ? '+' : ''}${(x * 100).toFixed(d)}%`;
  const delta = (now, before, lowerIsGood = false) => {
    if (before == null || !before) return '<div class="delta">No comparable previous period</div>';
    const d = (now - before) / before, up = d >= 0, good = lowerIsGood ? !up : up;
    return `<div class="delta"><span class="${good ? 'up-good' : 'up-bad'}">${ic(up ? 'up' : 'dn', 12)} ${pctS(d)}</span> ${periodMonths().prevLabel}</div>`;
  };
  const kpi = (label, value, foot = '') => `<div class="card kpi"><div class="label">${label}</div><div class="value">${value}</div>${foot}</div>`;
  const badge = (t = 'Synthetic data') => `<span class="badge" title="Values are simulated in the source workbook, not measured Bulk data">${t}</span>` +
    (UPLOADS.length && t === 'Synthetic data' ? ` <span class="badge up" title="${UPLOADS.map(u => u.name).join(', ')}">Includes uploaded data</span>` : '');
  const periodSelect = () => `<select data-bind="period">${Object.entries(PERIODS).map(([k, l]) => `<option value="${k}" ${k === state.period ? 'selected' : ''}>${l}</option>`).join('')}</select>`;
  const siteSelect = () => `<select data-bind="site">${SITES.map(s => `<option ${s === state.site ? 'selected' : ''}>${s}</option>`).join('')}</select>`;
  const rangeLabel = (ms = periodMonths().cur) => ms.length === 1 ? mLong(ms[0]) : `${mLong(ms[0])} – ${mLong(ms.at(-1))}`;
  const head = (title, sub, controls = '') => `<div class="head"><div><h1>${title} ${infoBtn('page:' + currentPage, 'title-info')}<span class="badge conf" title="Internal Bulk material. Do not share outside the company.">${ic('lock', 11)} Confidential</span></h1><div class="sub">${sub}</div></div><div class="controls">${controls}</div></div>`;
  const tabs = (key, items) => `<div class="tabs">${items.map(([k, l]) => `<button data-tab="${key}:${k}" class="${state[key] === k ? 'active' : ''}">${l}</button>`).join('')}</div>`;
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const note = t => `<div class="note">${ic('info', 13)} <span>${t}</span></div>`;

  /* ---------- chart helpers ---------- */
  const mk = (id, cfg) => { const el = document.getElementById(id); if (el) charts.push(new Chart(el, cfg)); };
  const opts = (extra = {}) => ({ responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom' } }, ...extra });
  const noGridX = { grid: { display: false } };

  const insightCard = (r, go = 'insights') => `<div class="insight" data-go="${go}" data-site="${r.site}"><span style="color:${C.amber}">${ic('warn')}</span><div><b>Higher cooling energy at ${r.site} – ${mLong(r.month)}</b>${pctS(r.residualPct, 0)} vs. expected (${r.coolingMWh.toFixed(0)} vs. ${r.expectedCoolingMWh.toFixed(0)} MWh). Not explained by IT load or outdoor temperature in the model.</div><span class="chev">${ic('chev')}</span></div>`;

  /* ---------- pages ---------- */
  const pages = {};

  function execSummary(a, flags) {
    const n01Last = rows([LAST], ['N01'])[0], cap = N01_CAPACITY ? N01_CAPACITY.value : null;
    const flagSites = [...new Set(flags.map(f => f.site))];
    return `<div class="card exec"><h3>${ic('book')} About this platform</h3>
      <p class="phone-only">Energy, cooling, cost and AI alerts for Bulk's data centres in one place.${flags.length ? ` <b>${flags.length} cooling flag(s)</b> need attention.` : ''}</p>
      <div class="desktop-only"><p>Bulk Energy Intelligence brings energy, cooling, cost, capacity and benchmark data for Bulk's data centres (${SITES.map(x => `${esc(x)} ${esc(SITE_INFO[x]?.location || '')}`).join(', ')}) together in one place. It helps operations, finance and management see how efficiently each site uses electricity, catch unusual cooling early, plan capacity, and compare Bulk with other Nordic operators. Hover over or click ${ic('info', 12)} on any card for a plain-language explanation.</p>
      <ul class="exec-points">
        <li><b>Efficiency:</b> average PUE ${a.pue.toFixed(3)} for ${rangeLabel()}, ${a.pue <= PUE_TARGET ? 'within' : `${((a.pue - PUE_TARGET) / PUE_TARGET * 100).toFixed(1)}% above`} the ${PUE_TARGET.toFixed(2)} design target.</li>
        <li><b>Energy and cost:</b> ${gwh(a.total)} used, about ${nokM(a.cost)} at market prices.</li>
        <li><b>Attention:</b> ${flags.length ? `${flags.length} cooling flag(s) at ${flagSites.join(', ')}. See AI Insights.` : 'no cooling flags in this period.'}${n01Last && cap ? ` N01 uses ${(n01Last.itLoadMW / cap * 100).toFixed(0)}% of its ${cap} MW capacity.` : ''}</li>
      </ul>
      <p class="small">Prototype for internal use. Operating figures are synthetic (from the project workbook); verified facts and sources are listed under Data &amp; Sources.</p></div></div>`;
  }
  pages.overview = () => {
    const a = cur(), p = prev(), ms = periodMonths().cur;
    const flags = flagged(ms), latestAny = flagged()[0];
    const excess = flags.reduce((t, r) => t + Math.max(r.residualMWh, 0), 0);
    const insights = flags.length
      ? flags.slice(0, 3).map(r => insightCard(r)).join('') +
        `<div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('bulb')}</span><div><b>Potential saving</b>Investigate cooling system performance at ${[...new Set(flags.map(f => f.site))].join(', ')}. Excess cooling vs. model: ${excess.toFixed(0)} MWh (${(excess / a.total * 100).toFixed(2)}% of total energy), about ${nokM(flags.reduce((t, r) => t + excessNOK(r), 0))} at the spot index.</div></div>`
      : `<div class="insight"><span style="color:${C.green}">${ic('check')}</span><div><b>No cooling review flags in this period</b>All site-months are within ±${THRESH * 100}% of expected cooling.</div></div>` +
        (latestAny ? insightCard(latestAny) : '');
    app.innerHTML = head('Overview', `Energy performance across all sites · ${rangeLabel()} ${badge()}`, periodSelect()) + execSummary(a, flags) + `
      <div class="grid g4">
        ${kpi('Average PUE', a.pue.toFixed(3), delta(a.pue, p?.pue, true))}
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total))}
        ${kpi('Cooling Energy', gwh(a.cool), `<div class="delta">${(a.coolShare * 100).toFixed(1)}% of total</div>`)}
        ${kpi('Estimated Energy Cost', nokM(a.cost), delta(a.cost, p?.cost, true))}
      </div>
      <div class="grid g21">
        <div class="card"><h3>Energy consumption by site</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3><span style="color:${C.purple}">${ic('spark')}</span> AI insights</h3>${insights}</div>
      </div>
      ${note('Cost = facility MWh × synthetic market price index (NOK/MWh). It is not Bulk\'s invoice, hedge or tariff. Source: workbook sheet "Sources &amp; Methods", M04.')}`;
    const by = SITES.map(s => cur([s]));
    mk('c1', { type: 'bar', data: { labels: SITES, datasets: [
      { label: 'IT energy', data: by.map(b => +(b.it / 1000).toFixed(2)), backgroundColor: C.purple },
      { label: 'Cooling energy', data: by.map(b => +(b.cool / 1000).toFixed(2)), backgroundColor: C.blue },
      { label: 'Other', data: by.map(b => +(b.other / 1000).toFixed(2)), backgroundColor: C.other }] },
      options: opts({ scales: { x: { stacked: true, ...noGridX }, y: { stacked: true, ticks: { callback: v => `${v} GWh` } } }, datasets: { bar: { maxBarThickness: 120 } } }) });
  };

  pages.sites = () => {
    const s = state.site, a = cur([s]), p = prev([s]), ms = periodMonths().cur;
    const mrows = ms.map(m => rows([m], [s])[0]).filter(Boolean);
    const lastRow = mrows.at(-1);
    const info = SITE_INFO[s] || {};
    let body = '';
    if (!lastRow) {
      app.innerHTML = head(`Site: ${s}`, `${esc(info.location)} · ${rangeLabel()}`, periodSelect()) + sitePicker() +
        `<div class="card"><div class="insight"><span style="color:${C.blue}">${ic('info')}</span><div><b>No data for ${s} in this period</b>Choose another period or site.</div></div></div>`;
      return;
    }
    if (state.siteTab === 'performance') {
      body = `<div class="grid g4">
        ${kpi('PUE', a.pue.toFixed(3), delta(a.pue, p?.pue, true))}
        ${kpi('IT Energy', gwh(a.it), delta(a.it, p?.it))}
        ${kpi('Cooling Energy', gwh(a.cool), delta(a.cool, p?.cool, true))}
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total, true))}</div>
        <div class="grid g2"><div class="card"><h3>PUE trend</h3><div class="chart"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3>Energy breakdown (${mLong(lastRow.month)})</h3><div class="donut-wrap"><div class="chart"><canvas id="c2"></canvas><div class="center-label"><b>${(lastRow.totalMWh / 1000).toFixed(1)}</b><span class="small">GWh</span></div></div><div class="legend">
          ${[['IT energy', lastRow.itMWh, C.purple], ['Cooling energy', lastRow.coolingMWh, C.blue], ['Other', lastRow.otherMWh, C.other]].map(([l, v, c]) => `<div style="display:flex"><span><span class="dot" style="background:${c}"></span>${l}</span><b>${(v / lastRow.totalMWh * 100).toFixed(0)}%</b></div>`).join('')}
        </div></div></div></div>
        ${note(`Dashed line: Bulk's new-build design PUE target (max. ${PUE_TARGET.toFixed(2)}), from the Annual Report 2025. It is a target, not a measured value.`)}`;
    } else if (state.siteTab === 'breakdown') {
      body = `<div class="card"><h3>Monthly energy breakdown (GWh)</h3><div class="chart tall"><canvas id="c3"></canvas></div></div>`;
    } else if (state.siteTab === 'cost') {
      body = `<div class="grid g4">${kpi('Estimated energy cost', nokM(a.cost), delta(a.cost, p?.cost, true))}${kpi('Avg. price', `${(a.price / 1000).toFixed(2)} NOK/kWh`, delta(a.price, p?.price, true))}${kpi('Price zone', esc(zone(s)), `<div class="delta">${esc(info.location)}</div>`)}${kpi('Avg. outdoor temp.', `${(mrows.reduce((t, r) => t + r.tempC, 0) / mrows.length).toFixed(1)} °C`, '<div class="delta">Synthetic weather</div>')}</div>
        <div class="card"><h3>Monthly cost (NOK M) and market price (NOK/MWh)</h3><div class="chart tall"><canvas id="c4"></canvas></div></div>`;
    } else {
      const fl = flagged(ms, [s]);
      body = `<div class="grid g21"><div class="card"><h3>Cooling residual vs. model (%)</h3><div class="chart tall"><canvas id="c5"></canvas></div></div>
        <div class="card"><h3><span style="color:${C.purple}">${ic('spark')}</span> AI analysis for ${s}</h3>
        ${fl.length ? fl.map(r => insightCard(r)).join('') : `<div class="insight"><span style="color:${C.green}">${ic('check')}</span><div><b>No review flags</b>Cooling at ${s} is within ±${THRESH * 100}% of the model for ${rangeLabel()}.</div></div>`}
        <div class="insight"><span style="color:${C.blue}">${ic('info')}</span><div><b>PUE vs. design target</b>PUE ${a.pue.toFixed(3)} vs. target max. ${PUE_TARGET.toFixed(2)}: ${a.pue <= PUE_TARGET ? 'within target' : `${((a.pue - PUE_TARGET) / PUE_TARGET * 100).toFixed(1)}% above target`}.</div></div></div></div>`;
    }
    app.innerHTML = head(`Site: ${s}`, `${esc(info.location)} · ${esc(zone(s))} · ${rangeLabel()} ${badge()}`, periodSelect()) + sitePicker() +
      tabs('siteTab', [['performance', 'Performance'], ['breakdown', 'Energy breakdown'], ['cost', 'Cost &amp; prices'], ['analysis', 'AI analysis']]) + body;
    const labels = mrows.map(r => mShort(r.month));
    const pues = mrows.map(r => +r.pue.toFixed(4));
    mk('c1', { type: 'line', data: { labels, datasets: [{ label: 'PUE', data: pues, borderColor: C.purple, backgroundColor: C.purple, pointRadius: 3 }] },
      options: opts({ plugins: { legend: { display: false }, overlays: { hlines: [{ y: PUE_TARGET, label: `Target ${PUE_TARGET.toFixed(2)}` }] } },
        scales: { x: noGridX, y: { suggestedMin: Math.min(PUE_TARGET, ...pues) - 0.02, suggestedMax: Math.max(...pues) + 0.02 } } }) });
    if (lastRow) mk('c2', { type: 'doughnut', data: { labels: ['IT energy', 'Cooling energy', 'Other'], datasets: [{ data: [lastRow.itMWh, lastRow.coolingMWh, lastRow.otherMWh].map(v => +(v / 1000).toFixed(2)), backgroundColor: [C.purple, C.blue, C.other], borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '72%', plugins: { legend: { display: false } } } });
    mk('c3', { type: 'bar', data: { labels, datasets: [
      { label: 'IT energy', data: mrows.map(r => +(r.itMWh / 1000).toFixed(2)), backgroundColor: C.purple },
      { label: 'Cooling energy', data: mrows.map(r => +(r.coolingMWh / 1000).toFixed(2)), backgroundColor: C.blue },
      { label: 'Other', data: mrows.map(r => +(r.otherMWh / 1000).toFixed(2)), backgroundColor: C.other }] },
      options: opts({ scales: { x: { stacked: true, ...noGridX }, y: { stacked: true } } }) });
    mk('c4', { data: { labels, datasets: [
      { type: 'line', label: 'Price (NOK/MWh)', data: mrows.map(r => r.priceNOKMWh), borderColor: C.amber, backgroundColor: C.amber, yAxisID: 'y1', pointRadius: 3 },
      { type: 'bar', label: 'Cost (NOK M)', data: mrows.map(r => +(r.costNOK / 1e6).toFixed(2)), backgroundColor: C.purple, yAxisID: 'y' }] },
      options: opts({ scales: { x: noGridX, y: { beginAtZero: true, title: { display: true, text: 'NOK M' } }, y1: { position: 'right', grid: { display: false }, beginAtZero: true, title: { display: true, text: 'NOK/MWh' } } } }) });
    mk('c5', { type: 'bar', data: { labels, datasets: [{ label: 'Residual %', data: mrows.map(r => +(r.residualPct * 100).toFixed(1)), backgroundColor: mrows.map(r => r.alert === 'REVIEW' ? C.amber : C.purple) }] },
      options: opts({ plugins: { legend: { display: false }, overlays: { hlines: [{ y: THRESH * 100, label: `Review threshold +${THRESH * 100}%`, color: C.amber }] } }, scales: { x: noGridX, y: { suggestedMax: 22, suggestedMin: -10, ticks: { callback: v => `${v}%` } } } }) });
  };

  pages.energy = () => {
    const a = cur(), p = prev(), ms = periodMonths().cur;
    app.innerHTML = head('Energy &amp; Cost', `Consumption and spend across all sites · ${rangeLabel()} ${badge()}`, periodSelect()) + `
      <div class="grid g4">
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total))}
        ${kpi('Total Cost', nokM(a.cost), delta(a.cost, p?.cost, true))}
        ${kpi('Avg. price', `${(a.price / 1000).toFixed(2)} NOK/kWh`, delta(a.price, p?.price, true))}
        ${kpi('Cooling share', `${(a.coolShare * 100).toFixed(1)}%`, delta(a.coolShare, p?.coolShare, true))}
      </div>
      <div class="card" style="margin-bottom:12px"><h3>Monthly energy (GWh) and cost (NOK M)</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
      <div class="card" style="margin-bottom:12px"><h3>Cost by site</h3><table><thead><tr><th>Site</th><th>Location</th><th>Price zone</th><th>Energy</th><th>PUE</th><th>Avg. price</th><th>Est. cost</th></tr></thead><tbody>
        ${SITES.map(s => { const b = cur([s]); return `<tr><td><b>${s}</b></td><td>${esc(SITE_INFO[s]?.location)}</td><td>${esc(zone(s))}</td><td>${gwh(b.total)}</td><td>${b.pue.toFixed(3)}</td><td>${(b.price / 1000).toFixed(2)} NOK/kWh</td><td>NOK ${(b.cost / 1e6).toFixed(2)}M</td></tr>`; }).join('')}
      </tbody></table></div>
      ${note('The workbook has no renewable-share data, so that KPI from the design is replaced by cooling share. Prices are a synthetic monthly index per price zone, not Nord Pool history.')}`;
    const mr = ms.map(m => agg(rows([m])));
    mk('c1', { data: { labels: ms.map(mShort), datasets: [
      { type: 'line', label: 'Cost (NOK M)', data: mr.map(r => +(r.cost / 1e6).toFixed(2)), borderColor: C.amber, backgroundColor: C.amber, pointRadius: 3, yAxisID: 'y1' },
      { type: 'bar', label: 'Energy (GWh)', data: mr.map(r => +(r.total / 1000).toFixed(2)), backgroundColor: C.purple, yAxisID: 'y' }] },
      options: opts({ scales: { x: noGridX, y: { beginAtZero: true, title: { display: true, text: 'GWh' } }, y1: { position: 'right', beginAtZero: true, grid: { display: false }, title: { display: true, text: 'NOK M' } } } }) });
  };

  const ACTIONS = ['Check cooling system performance (setpoints, pumps, condenser)', 'Review recent maintenance events', 'Compare with other sites', 'Validate the reading against permissioned metering before acting'];
  pages.insights = () => {
    const s = state.site, ms = periodMonths().cur;
    const siteRows = rows(MONTHS, [s]);
    const inPeriod = flagged(ms, [s]);
    const main = inPeriod[0] || flagged(MONTHS, [s])[0];
    const done = store.get('actions', {});
    let banner, impact = '';
    if (main) {
      const win = (inPeriod.length ? inPeriod : [main]);
      const exMWh = win.reduce((t, r) => t + Math.max(r.residualMWh, 0), 0);
      const exNOK = win.reduce((t, r) => t + excessNOK(r), 0);
      const totalCool = agg(win).cool;
      banner = `<div class="banner"><span style="color:${C.amber}">${ic('warn', 26)}</span><div><b>Unusual increase in cooling energy</b><p>Cooling energy at ${s} was ${(main.residualPct * 100).toFixed(0)}% higher than expected in ${mLong(main.month)}, while IT load and outdoor temperature do not explain it (model residual above the ${THRESH * 100}% threshold).${inPeriod.length ? '' : ' This flag is outside the selected period.'}</p></div><button class="btn" data-action="details" data-month="${main.month}">View details ${ic('arrow', 13)}</button></div>`;
      const key = s + main.month;
      impact = `<div class="grid g21"><div class="card"><h3>Actual vs. expected cooling energy (MWh)</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3><span style="color:${C.purple}">${ic('target')}</span> Suggested actions</h3>
          ${ACTIONS.map((t, i) => `<label class="check ${done[key + i] ? 'done' : ''}"><input type="checkbox" data-check="${key + i}" ${done[key + i] ? 'checked' : ''}><span>${t}</span></label>`).join('')}
          <div class="small" style="margin-top:14px">${ACTIONS.filter((_, i) => done[key + i]).length} of ${ACTIONS.length} actions completed</div></div></div>
        <div class="card" style="margin-bottom:12px"><h3><span style="color:${C.green}">${ic('leaf')}</span> Estimated impact (${win.slice().reverse().map(r => mShort(r.month)).join(', ')})</h3><div class="grid g3" style="margin:0">
          <div class="impact"><div class="label">${ic('link', 12)} Potential energy saving</div><div class="value">−${(exMWh / 1000).toFixed(2)} GWh</div></div>
          <div class="impact"><div class="label">${ic('link', 12)} Estimated cost saving</div><div class="value">NOK ${(exNOK / 1e6).toFixed(2)}M</div></div>
          <div class="impact"><div class="label">${ic('link', 12)} Share of cooling energy</div><div class="value">${(exMWh / totalCool * 100).toFixed(1)}%</div></div></div></div>`;
    } else {
      banner = `<div class="banner ok"><span style="color:${C.green}">${ic('check', 26)}</span><div><b>No unusual cooling at ${s}</b><p>All months are within ±${THRESH * 100}% of the expected-cooling model.</p></div></div>`;
      impact = `<div class="card" style="margin-bottom:12px"><h3>Actual vs. expected cooling energy (MWh)</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>`;
    }
    const others = flagged().filter(r => r.site !== s);
    app.innerHTML = head('AI Insights', `AI-powered analysis of energy use, efficiency and capacity ${badge()}`, periodSelect()) + sitePicker() + banner + impact +
      (others.length ? `<div class="card" style="margin-bottom:12px"><h3>Flags at other sites</h3>${others.map(r => insightCard(r)).join('')}</div>` : '') +
      note(`How the "AI model" works: expected cooling = IT MWh × (0.115 + 0.0055 × max(temperature − 8, 0) + site adjustment). It is a transparent heuristic, not a trained machine-learning model, and the anomalies were injected into synthetic data (workbook M05/M06). Have site engineers confirm any finding before acting.`);
    const labels = siteRows.map(r => mShort(r.month));
    // Merge consecutive flagged months into one shaded box with a combined label.
    const boxes = [];
    siteRows.forEach((r, i) => {
      if (r.alert !== 'REVIEW') return;
      const prevBox = boxes.at(-1);
      if (prevBox && prevBox.to === i - 1) { prevBox.to = i; prevBox.label += ` / ${pctS(r.residualPct, 0)}`; }
      else boxes.push({ from: i, to: i, label: pctS(r.residualPct, 0) });
    });
    const sel = siteRows.map(r => ms.includes(r.month));
    const first = sel.indexOf(true), last = sel.lastIndexOf(true);
    if (first > 0 || last < siteRows.length - 1) boxes.unshift({ from: first, to: last, color: 'rgba(255,255,255,0.035)' });
    mk('c1', { type: 'line', data: { labels, datasets: [
      { label: 'Actual', data: siteRows.map(r => r.coolingMWh), borderColor: C.purple, backgroundColor: C.purple, pointRadius: 2.5, tension: 0.25, borderWidth: 2 },
      { label: 'Expected (model)', data: siteRows.map(r => +r.expectedCoolingMWh.toFixed(1)), borderColor: '#cfd2e0', borderDash: [4, 4], pointRadius: 0, tension: 0.3, borderWidth: 1.2 }] },
      options: opts({ scales: { x: { ...noGridX, ticks: { maxTicksLimit: 11 } } }, plugins: { legend: { position: 'bottom' }, overlays: { boxes } } }) });
  };

  pages.capacity = () => {
    const lastRows = rows([LAST]);
    const n01 = lastRows.find(r => r.site === 'N01');
    const cap = N01_CAPACITY ? N01_CAPACITY.value : null;
    const totalLoad = lastRows.reduce((t, r) => t + r.itLoadMW, 0);
    const card = r => {
      const isN01 = r.site === 'N01' && cap;
      const u = isN01 ? r.itLoadMW / cap * 100 : null;
      const peak = Math.max(...rows(MONTHS, [r.site]).map(x => x.itLoadMW));
      return `<div class="card"><h3>${r.site} · ${esc(SITE_INFO[r.site]?.location)}</h3>
        ${isN01 ? `<div class="big">${u.toFixed(0)}%</div><div class="small">${r.itLoadMW.toFixed(1)} of ${cap} MW reported IT capacity</div><div class="bar"><div style="width:${u}%;${u >= 85 ? `background:${C.red}` : ''}"></div></div>
          <div class="small" style="color:#c9ccda">${u >= 85 ? 'Critical – expansion required' : u >= 75 ? 'Approaching limit – plan expansion' : 'Healthy headroom'}</div>`
          : `<div class="big">${r.itLoadMW.toFixed(1)} MW</div><div class="small">Average IT load, ${mLong(LAST)} (12-month peak ${peak.toFixed(1)} MW)</div><div class="bar"><div style="width:0"></div></div><div class="small" style="color:#c9ccda">Installed capacity not verified – utilisation not shown</div>`}
      </div>`;
    };
    app.innerHTML = head('Capacity', `Power utilisation and available headroom · ${mLong(LAST)}`) + `
      <div class="grid g3">
        ${kpi('Installed capacity (N01)', cap ? `${cap} MW` : 'n/a', `<div class="delta">Reported IT capacity, ${N01_CAPACITY?.period || ''} · planned ${N01_PLANNED?.value || '?'} MW by 2031</div>`)}
        ${kpi('Utilised (all sites)', `${totalLoad.toFixed(1)} MW`, `<div class="delta">Average IT load, ${mLong(LAST)} ${badge()}</div>`)}
        ${kpi('Available headroom (N01)', cap && n01 ? `${(cap - n01.itLoadMW).toFixed(1)} MW` : 'n/a', `<div class="delta">${cap && n01 ? `${(n01.itLoadMW / cap * 100).toFixed(0)}% of N01 capacity used` : ''}</div>`)}
      </div>
      <div class="grid g3">${SITES.map(s => lastRows.find(r => r.site === s)).filter(Boolean).map(card).join('')}</div>
      <div class="card" style="margin-bottom:12px"><h3>Average IT load by site (MW)</h3><div class="chart"><canvas id="c1"></canvas></div></div>
      ${note(`N01's ${cap} MW is the operating IT capacity reported for 2025 (Menon Economics, Aug 2026). It is capacity, not observed load. The workbook has no verified capacity for OS-IX or DK01, so no utilisation is shown for them.`)}`;
    const colors = { 'OS-IX': C.blue, 'N01': C.purple, 'DK01': C.green };
    mk('c1', { type: 'line', data: { labels: MONTHS.map(mShort), datasets: SITES.map(s => ({ label: s, spanGaps: true, data: MONTHS.map(m => (rows([m], [s])[0] || {}).itLoadMW ?? null), borderColor: colors[s], backgroundColor: colors[s], pointRadius: 0, borderWidth: 2 })) },
      options: opts({ scales: { x: noGridX, y: { beginAtZero: true } }, plugins: { legend: { position: 'bottom' }, overlays: { hlines: cap ? [{ y: cap, label: `N01 capacity ${cap} MW` }] : [] } } }) });
  };

  /* ---------- benchmarking: clickable competitor comparison ---------- */
  const SEL_COLORS = ['#f6b93b', '#4fa3f0', '#ff7ab6', '#b39dff'];
  const MAX_SEL = 4;
  const compKey = b => `${b.operator} · ${b.facility}`;
  const fp = b => b.pue.toFixed(b.merged ? 3 : 2) + (b.parts ? `<span class="small"> avg. (${b.parts.map(p => `${esc(p.facility.replace(/ campus$/i, ''))} ${p.pue.toFixed(2)}`).join(', ')})</span>` : '');
  const competitors = () => D.benchmarks.filter(b => !/^Bulk/i.test(b.operator) && !/TARGET/i.test(b.status || ''));
  const peerNames = () => [...new Set(D.peerMonthly.map(r => r.peer))];
  // Link a synthetic monthly series (e.g. "Verne Iceland") to its reported benchmark row.
  const peerFor = b => {
    if (b.merged) return peerNames().find(n => n === b.operator) || null;
    const word = String(b.facility || '').split(' ')[0].toLowerCase();
    return peerNames().find(n => n.toLowerCase() === `${b.operator} ${word}`.toLowerCase())
      || (/portfolio/i.test(b.facility || '') ? peerNames().find(n => n === b.operator) : null) || null;
  };
  const benchForPeer = peer => competitors().find(b => peerFor(b) === peer);
  state.benchSel = store.get('benchSel', []);
  const selected = () => state.benchSel.filter(k => competitors().some(b => compKey(b) === k));
  const selColor = k => SEL_COLORS[selected().indexOf(k)];
  function toggleBench(k) {
    let s = selected();
    if (s.includes(k)) s = s.filter(x => x !== k);
    else { if (s.length >= MAX_SEL) s.shift(); s.push(k); }
    state.benchSel = s; store.set('benchSel', s);
    setTimeout(render); // let Chart.js finish its click handler first
  }
  const peerStats = peer => {
    const r = D.peerMonthly.filter(x => x.peer === peer && typeof x.pue === 'number');
    if (!r.length) return null;
    const months = r.map(x => x.month);
    const bulk = months.map(m => agg(rows([m])).pue);
    return {
      months, n: r.length,
      avg: r.reduce((t, x) => t + x.totalMWh, 0) / r.reduce((t, x) => t + x.itMWh, 0),
      min: Math.min(...r.map(x => x.pue)), max: Math.max(...r.map(x => x.pue)),
      bulkLower: r.filter((x, i) => bulk[i] && bulk[i] < x.pue).length,
      bulkAvg: agg(rows(months)).pue, bulkMin: Math.min(...bulk.filter(Boolean)), bulkMax: Math.max(...bulk.filter(Boolean)),
    };
  };

  function comparisonCard(b2025) {
    const sel = selected().map(k => competitors().find(b => compKey(b) === k));
    if (!sel.length) return `<div class="card compare-empty" style="margin-bottom:12px">${ic('target', 18)}<span><b>Click a competitor</b> in the chart, the list or the chart legend to compare it side by side with Bulk. You can pick up to ${MAX_SEL}.</span></div>`;
    const stats = sel.map(b => peerFor(b) ? peerStats(peerFor(b)) : null);
    const bulkStats = stats.find(Boolean);
    const diff = b => {
      if (typeof b.pue !== 'number') return '<span class="small">No published PUE</span>';
      const d = b.pue - b2025.pue, rel = d / b2025.pue;
      return `<span class="${d > 0 ? 'up-good' : d < 0 ? 'up-bad' : ''}">${d >= 0 ? '+' : ''}${d.toFixed(3)} (${pctS(rel)})</span><div class="small">${d > 0 ? 'Bulk lower (more efficient)' : d < 0 ? 'Competitor lower' : 'Same'}</div>`;
    };
    const vsTarget = p => typeof p === 'number' ? `${p - PUE_TARGET >= 0 ? '+' : ''}${(p - PUE_TARGET).toFixed(3)}` : '—';
    const atScale = b => {
      if (typeof b.pue !== 'number') return '—';
      const mwh = b2025.it * (b.pue - b2025.pue), nok = mwh * b2025.price;
      return `${mwh >= 0 ? '+' : '−'}${Math.abs(mwh / 1000).toFixed(1)} GWh<div class="small">${mwh >= 0 ? '+' : '−'}NOK ${Math.abs(nok / 1e6).toFixed(1)}M at Bulk's avg. price</div>`;
    };
    const link = u => /^https?:/.test(u || '') ? `<a href="${esc(u)}" target="_blank" rel="noopener">Report ${ic('ext', 11)}</a>` : esc(u || '—');
    const rowsDef = [
      ['Reported PUE', `<b>${b2025.pue.toFixed(3)}</b><div class="small">2025, synthetic</div>`, b => typeof b.pue === 'number' ? `<b>${fp(b)}</b><div class="small">${esc(b.year ?? '')}</div>` : '<span class="small">Not published</span>'],
      ['Difference vs. Bulk', '—', diff],
      [`vs. Bulk design target ${PUE_TARGET.toFixed(2)}`, vsTarget(b2025.pue), b => vsTarget(b.pue)],
      ['If Bulk ran at this PUE (2025)', '—', atScale],
      ['Monthly scenario avg. (synthetic)', bulkStats ? bulkStats.bulkAvg.toFixed(3) : '—', (b, i) => stats[i] ? stats[i].avg.toFixed(3) : '<span class="small">No series</span>'],
      ['Monthly range', bulkStats ? `${bulkStats.bulkMin.toFixed(3)} – ${bulkStats.bulkMax.toFixed(3)}` : '—', (b, i) => stats[i] ? `${stats[i].min.toFixed(3)} – ${stats[i].max.toFixed(3)}` : '—'],
      ['Months where Bulk is lower', '—', (b, i) => stats[i] ? `${stats[i].bulkLower} of ${stats[i].n}` : '—'],
      ['Reported energy', `${(b2025.total / 1000).toFixed(1)} GWh<div class="small">2025 facility, synthetic</div>`, b => typeof b.energyMWh === 'number' ? `${(b.energyMWh / 1000).toFixed(1)} GWh` : '<span class="small">Not available</span>'],
      ['Energy scope', 'Facility electricity', b => esc(b.energyScope || '—')],
      ['PUE type', 'Energy-weighted annual (total / IT)', b => esc(b.pueType || '—')],
      ['Status', '<span class="badge">SYNTHETIC</span>', b => esc(b.status || '—')],
      ['Relevance to Bulk', 'Subject of the dashboard', b => esc(b.relevance || '—')],
      ['Caution', 'Simulated operating data', b => `<span class="small">${esc(b.cautions || '—')}</span>`],
      ['Source', 'Workbook (M01)', b => link(b.source)],
    ];
    const headCells = sel.map(b => `<th><span class="chip" style="background:${selColor(compKey(b))}"></span>${esc(b.operator)}<div class="small">${esc(b.facility)}</div><button class="x" data-bench="${esc(compKey(b))}" title="Remove from comparison">×</button></th>`).join('');
    return `<div class="card compare" style="margin-bottom:12px"><h3>${ic('target')} Side-by-side comparison <button class="btn ghost sm" data-action="bench-clear" style="margin-left:auto">Clear</button></h3>
      <div class="table-wrap"><table class="cmp"><thead><tr><th></th><th><span class="chip" style="background:${C.green}"></span>Bulk<div class="small">All sites</div></th>${headCells}</tr></thead><tbody>
      ${rowsDef.map(([label, bulk, f]) => `<tr><td class="small">${label}</td><td>${bulk}</td>${sel.map((b, i) => `<td>${f(b, i)}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>
      ${stats.some(Boolean) ? `<h3 style="margin-top:16px">Monthly PUE: Bulk vs. selected (synthetic scenarios)</h3><div class="chart"><canvas id="c3"></canvas></div>` : ''}
      ${note('"If Bulk ran at this PUE" = Bulk\'s 2025 IT energy × (competitor PUE − Bulk PUE). It shows what the PUE gap means in energy and cost at Bulk\'s scale. It is an illustration, because years, scopes and methods differ.')}
    </div>`;
  }

  pages.benchmarking = () => {
    const b2025 = agg(rows(MONTHS.filter(m => m.startsWith('2025'))));
    const comps = competitors();
    const numeric = comps.filter(b => typeof b.pue === 'number' && /^VERIFIED/.test(b.status || ''));
    // Independent peers: skip EcoDataCenter Falun (same operator as portfolio row).
    const peers = numeric.filter(b => !/Falun/.test(b.facility));
    const mkt = peers.reduce((t, b) => t + b.pue, 0) / peers.length;
    const names = peerNames();
    const peerMonths = [...new Set(D.peerMonthly.map(r => r.month))].sort();
    const sel = selected(), anySel = sel.length > 0;
    const rankRow = (b, value, label) => {
      const k = compKey(b), on = sel.includes(k);
      return `<button class="rank pick ${on ? 'sel' : ''}" data-bench="${esc(k)}" style="${on ? `--sel:${selColor(k)}` : ''}" aria-pressed="${on}"><span>${label}</span><b>${value}</b><i class="tick">${on ? ic('check', 13) : '+'}</i></button>`;
    };
    let left, right;
    if (state.benchTab === 'reported') {
      left = 'Reported annual PUE (public reports)';
      right = `<div class="rank me"><span>Bulk 2025 (synthetic)</span><b>${b2025.pue.toFixed(3)}</b></div>` +
        comps.map(b => rankRow(b, typeof b.pue === 'number' ? fp(b) : '—', `${esc(b.operator)} · ${esc(b.facility)} <span class="small">${b.year ? `(${b.year})` : '(no verified PUE)'}</span>`)).join('') +
        `<div class="rank"><span>Peer average (excl. Falun)</span><b>${mkt.toFixed(2)}</b></div>`;
    } else if (state.benchTab === 'monthly') {
      left = 'Monthly PUE scenarios 2025';
      right = `<div class="rank me"><span>Bulk (all sites)</span><b>${b2025.pue.toFixed(3)}</b></div>` +
        names.map(n => { const b = benchForPeer(n), st = peerStats(n); return b ? rankRow(b, st.avg.toFixed(3), esc(n)) : `<div class="rank"><span>${esc(n)}</span><b>${st.avg.toFixed(3)}</b></div>`; }).join('');
    } else {
      left = 'Market price index by zone (NOK/MWh)';
      right = SITES.map(s => { const r = rows(MONTHS, [s]); const avg = r.reduce((t, x) => t + x.priceNOKMWh, 0) / r.length; return `<div class="rank"><span>${esc(zone(s))} (${s})</span><b>${avg.toFixed(0)}</b></div>`; }).join('');
    }
    const n01 = rows(MONTHS, ['N01']);
    const hi = n01.length ? n01.reduce((m, r) => r.priceNOKMWh > m.priceNOKMWh ? r : m, n01[0]) : null;
    const avgP = n01.length ? n01.reduce((t, r) => t + r.priceNOKMWh, 0) / n01.length : 0;
    const selTake = sel.map(k => comps.find(b => compKey(b) === k)).filter(b => typeof b.pue === 'number')
      .map(b => `${esc(b.operator)} (${esc(b.facility)}) reports ${fp(b)}, which is ${Math.abs((b.pue - b2025.pue) / b2025.pue * 100).toFixed(1)}% ${b.pue > b2025.pue ? 'higher' : 'lower'} than Bulk.`).join(' ');
    const showCompare = state.benchTab !== 'price';
    app.innerHTML = head('Benchmarking', `Compare Bulk's sites with selected competitors and market context`, '') +
      tabs('benchTab', [['reported', 'Reported PUE (verified)'], ['monthly', 'Monthly scenarios (synthetic)'], ['price', 'Electricity price']]) + `
      <div class="grid g21"><div class="card"><h3>${left}${showCompare ? '<span class="small" style="margin-left:auto;font-weight:400">Click to compare</span>' : ''}</h3><div class="chart tall"><canvas id="c1"></canvas></div></div><div class="card"><h3>${state.benchTab === 'price' ? 'Average price (all months)' : 'Average PUE'}</h3>${right}</div></div>
      ${showCompare ? comparisonCard(b2025) : ''}
      <div class="grid g21"><div class="card"><h3>Electricity price N01 (NOK/MWh)</h3><div class="chart"><canvas id="c2"></canvas></div></div>
      <div class="card"><h3><span style="color:${C.green}">${ic('info')}</span> Key takeaway</h3><p class="takeaway">
        Bulk's 2025 PUE in the synthetic data (${b2025.pue.toFixed(3)}) is ${((mkt - b2025.pue) / mkt * 100).toFixed(1)}% lower than the average reported by Nordic peers (${mkt.toFixed(2)}). ${selTake}
        Because Bulk's figure is simulated and the peers use different years, scopes and methods, this is an illustration, not a ranking.
        ${hi ? `N01's price index peaked in ${mLong(hi.month)} (${hi.priceNOKMWh.toFixed(0)} NOK/MWh), ${((hi.priceNOKMWh - avgP) / avgP * 100).toFixed(0)}% above its average.` : ''}</p></div></div>
      ${state.benchTab === 'reported' ? `<div class="card" style="margin-bottom:12px"><h3>Comparability notes</h3><table><thead><tr><th>Operator</th><th>Scope</th><th>Status</th><th>Caution</th><th>Source</th></tr></thead><tbody>
        ${D.benchmarks.map(b => `<tr class="${sel.includes(compKey(b)) ? 'row-sel' : ''}"><td><b>${esc(b.operator)}</b><div class="small">${esc(b.facility)}</div></td><td>${esc(b.pueType)}</td><td>${esc(b.status)}</td><td>${esc(b.cautions)}</td><td>${/^https?:/.test(b.source || '') ? `<a href="${esc(b.source)}" target="_blank" rel="noopener">Link ${ic('ext', 11)}</a>` : esc(b.source)}</td></tr>`).join('')}
      </tbody></table></div>` : ''}
      ${state.benchTab === 'monthly' ? note('Monthly peer values are synthetic scenarios anchored only to each peer\'s published annual PUE. They are not observed peer operations.') : ''}`;

    const pointer = (evt, els) => { evt.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
    if (state.benchTab === 'reported') {
      const labels = ['Bulk 2025*', ...numeric.map(b => `${b.operator}${b.merged || /portfolio/i.test(b.facility) ? '' : ' ' + b.facility.split(' ')[0]}`)];
      const barColor = b => sel.includes(compKey(b)) ? selColor(compKey(b)) : anySel ? '#33374f' : '#4a4f6a';
      mk('c1', { type: 'bar', data: { labels, datasets: [{ data: [+b2025.pue.toFixed(3), ...numeric.map(b => b.pue)], backgroundColor: [C.green, ...numeric.map(barColor)], maxBarThickness: 60 }] },
        options: opts({
          onHover: pointer,
          onClick: (evt, els) => { const i = els[0]?.index; if (i > 0) toggleBench(compKey(numeric[i - 1])); },
          interaction: { mode: 'nearest', intersect: false, axis: 'x' },
          plugins: { legend: { display: false }, overlays: { hlines: [{ y: PUE_TARGET, label: `Bulk design target ${PUE_TARGET.toFixed(2)}` }] }, tooltip: { callbacks: { afterLabel: c => c.dataIndex === 0 ? 'Synthetic, not reported' : `${numeric[c.dataIndex - 1].year} · ${numeric[c.dataIndex - 1].pueType}\nClick to ${sel.includes(compKey(numeric[c.dataIndex - 1])) ? 'remove from' : 'add to'} comparison` } } },
          scales: { x: noGridX, y: { min: 1.0, max: 1.45 } } }) });
    } else if (state.benchTab === 'monthly') {
      const bulkM = peerMonths.map(m => +agg(rows([m])).pue.toFixed(4));
      const shades = ['#e6e7ee', '#b8bbd0', '#8f93ad', '#6d7190', '#55597a'];
      const ds = names.map((n, i) => {
        const b = benchForPeer(n), k = b && compKey(b), on = k && sel.includes(k);
        return { label: n, data: peerMonths.map(m => (D.peerMonthly.find(r => r.peer === n && r.month === m) || {}).pue),
          borderColor: on ? selColor(k) : anySel ? '#3a3e57' : shades[i % shades.length], backgroundColor: on ? selColor(k) : 'transparent',
          borderDash: !on && i % 2 ? [5, 4] : [], pointRadius: on ? 2.5 : 0, borderWidth: on ? 2.5 : 1 };
      });
      mk('c1', { type: 'line', data: { labels: peerMonths.map(mShort), datasets: [
        { label: 'Bulk (avg.)', data: bulkM, borderColor: C.green, backgroundColor: C.green, borderWidth: 2.5, pointRadius: 3 }, ...ds] },
        options: opts({
          onHover: (evt, els, chart) => { evt.native.target.style.cursor = chart.getElementsAtEventForMode(evt, 'nearest', { intersect: true }, false).some(e => e.datasetIndex > 0) ? 'pointer' : 'default'; },
          onClick: (evt, _e, chart) => { const hit = chart.getElementsAtEventForMode(evt, 'nearest', { intersect: false }, false)[0]; if (hit && hit.datasetIndex > 0) { const b = benchForPeer(names[hit.datasetIndex - 1]); if (b) toggleBench(compKey(b)); } },
          plugins: { legend: { position: 'bottom', onClick: (e, item) => { if (item.datasetIndex > 0) { const b = benchForPeer(names[item.datasetIndex - 1]); if (b) toggleBench(compKey(b)); } }, onHover: e => { e.native.target.style.cursor = 'pointer'; } } },
          scales: { x: noGridX, y: { suggestedMin: 1.15, suggestedMax: 1.42 } } }) });
    } else {
      const palette = [C.blue, '#2f9e5f', C.amber, '#ff7ab6', '#b39dff'];
      mk('c1', { type: 'line', data: { labels: MONTHS.map(mShort), datasets: SITES.map((s, i) => ({ label: zone(s) || s, spanGaps: true, data: MONTHS.map(m => (rows([m], [s])[0] || {}).priceNOKMWh ?? null), borderColor: palette[i % palette.length], backgroundColor: palette[i % palette.length], pointRadius: 0, borderWidth: 2 })) },
        options: opts({ scales: { x: noGridX, y: { beginAtZero: true } } }) });
    }
    // Comparison chart: Bulk vs. selected peers (monthly synthetic scenarios).
    const selPeers = sel.map(k => comps.find(b => compKey(b) === k)).filter(b => b && peerFor(b));
    if (selPeers.length) {
      mk('c3', { type: 'line', data: { labels: peerMonths.map(mShort), datasets: [
        { label: 'Bulk (avg.)', data: peerMonths.map(m => +agg(rows([m])).pue.toFixed(4)), borderColor: C.green, backgroundColor: C.green, borderWidth: 2.5, pointRadius: 2.5 },
        ...selPeers.map(b => ({ label: `${b.operator} · ${b.facility}`, data: peerMonths.map(m => (D.peerMonthly.find(r => r.peer === peerFor(b) && r.month === m) || {}).pue), borderColor: selColor(compKey(b)), backgroundColor: selColor(compKey(b)), borderWidth: 2, pointRadius: 2 }))] },
        options: opts({ plugins: { legend: { position: 'bottom' }, overlays: { hlines: [{ y: PUE_TARGET, label: `Bulk design target ${PUE_TARGET.toFixed(2)}` }] } }, scales: { x: noGridX, y: { suggestedMin: 1.15, suggestedMax: 1.42 } } }) });
    }
    if (n01.length) mk('c2', { type: 'line', data: { labels: n01.map(r => mShort(r.month)), datasets: [{ label: 'NOK/MWh', data: n01.map(r => r.priceNOKMWh), borderColor: '#2f9e5f', backgroundColor: 'rgba(47,158,95,0.35)', fill: true, pointRadius: 0, borderWidth: 1.5 }] },
      options: opts({ plugins: { legend: { display: false } }, scales: { x: { ...noGridX, ticks: { maxTicksLimit: 11 } }, y: { beginAtZero: true } } }) });
  };

  /* ---------- reports ---------- */
  const today = new Date().toISOString().slice(0, 10);
  const nextDay = m => { const [y, mo] = m.split('-').map(Number); return mo === 12 ? `${y + 1}-01-01` : `${y}-${String(mo + 1).padStart(2, '0')}-01`; };
  const BUILT_IN = [
    { title: `Monthly energy report – ${mLong(LAST)}`, date: nextDay(LAST), type: 'PDF', kind: 'monthly', month: LAST },
    { title: `${LAST.slice(0, 4)} year-to-date sustainability summary`, date: nextDay(LAST), type: 'PDF', kind: 'ytd' },
    ...flagged().slice(0, 2).map(r => ({ title: `${r.site} cooling anomaly analysis – ${mLong(r.month)}`, date: nextDay(r.month), type: 'PDF', kind: 'anomaly', site: r.site, month: r.month })),
    { title: 'Capacity and IT load data', date: nextDay(LAST), type: 'XLSX', kind: 'capacity' },
    { title: 'Annual energy report 2025', date: '2026-01-15', type: 'PDF', kind: 'annual' },
    { title: `Monthly energy report – ${mLong(MONTHS.at(-2))}`, date: nextDay(MONTHS.at(-2)), type: 'PDF', kind: 'monthly', month: MONTHS.at(-2) },
  ];
  const allReports = () => [...store.get('reports', []), ...BUILT_IN];
  const fmtDate = d => { const x = new Date(d + 'T00:00:00'); return `${x.getDate()} ${MON[x.getMonth()]} ${x.getFullYear()}`; };
  pages.reports = () => {
    const ups = window.BulkUpload ? window.BulkUpload.list() : [];
    const upload = `<div class="card upload-card" style="margin-bottom:12px"><h3>${ic('up', 15)} Upload a report to update the dashboard</h3>
        <label class="dropzone" id="dropzone">
          <input type="file" id="upload-input" accept=".xlsx,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>
          <div>${ic('down', 22)}</div>
          <b>Drop an Excel (.xlsx) or CSV file here, or click to choose</b>
          <span class="small">Use the Bulk data workbook format, or a sheet with columns such as Month, Site, IT energy (MWh), Cooling (MWh), Other (MWh), Market price NOK/MWh. You'll see a preview before anything changes.</span>
        </label>
        <div class="row-end" style="justify-content:flex-start;margin-top:10px"><button class="btn ghost" data-action="template">${ic('file', 13)} Download template (current data as CSV)</button></div>
        ${ups.length ? `<h3 style="margin-top:16px">Uploaded data (${ups.length})</h3>${ups.slice().reverse().map(u => `<div class="report"><span class="ic">${ic('db', 18)}</span><div><b>${esc(u.name)}</b><div class="small">${fmtDate(u.date.slice(0, 10))} · ${u.mode === 'replace' ? 'replaced monthly data' : 'merged'} · ${u.summary.monthlyRows} rows (${u.summary.added} new, ${u.summary.updated} updated)${u.summary.range ? ` · ${mShort(u.summary.range[0])} – ${mShort(u.summary.range[1])}` : ''}${u.summary.tables.length ? ` · ${u.summary.tables.length} other table(s)` : ''}</div></div><button class="btn ghost" data-action="remove-upload" data-id="${u.id}">Remove</button></div>`).join('')}
          <div class="row-end" style="justify-content:flex-start"><button class="btn ghost" data-action="reset-uploads">Restore original data</button></div>` : ''}
      </div>`;
    app.innerHTML = head('Reports', 'Download generated reports or upload new data', '<button class="btn" data-action="new-report">New report</button>') + upload +
      `<div class="card"><h3>${ic('file')} Available reports</h3>${allReports().map((r, i) => `<div class="report"><span class="ic">${ic('file', 18)}</span><div><b>${esc(r.title)}</b><div class="small">${fmtDate(r.date)} · ${r.type === 'XLSX' ? 'XLSX (CSV)' : r.type}</div></div><button class="btn ghost" data-download="${i}">${ic('down', 13)} Download</button></div>`).join('')}</div>
      ${note('PDF reports open in a print view. Choose "Save as PDF" in the print dialog. Data exports download as CSV, which opens in Excel. Uploaded data is stored only in this browser.')}`;
    const dz = document.getElementById('dropzone'), input = document.getElementById('upload-input');
    input.addEventListener('change', () => { if (input.files[0]) handleUpload(input.files[0]); input.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
    dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleUpload(f); });
  };

  const MONTHLY_HEADER = ['Month', 'Site', 'Zone', 'IT avg load (MW)', 'IT energy (MWh)', 'Cooling (MWh)', 'Other (MWh)', 'Total (MWh)', 'PUE', 'Outdoor °C', 'Price NOK/MWh', 'Spot-index cost NOK', 'Expected cooling (MWh)', 'Residual %', 'Alert', 'Provenance'];
  const monthlyCsvRows = () => D.monthly.map(m => [m.month, m.site, m.zone, m.itLoadMW, m.itMWh, m.coolingMWh, m.otherMWh, m.totalMWh, m.pue, m.tempC, m.priceNOKMWh, m.costNOK, m.expectedCoolingMWh, m.residualPct, m.alert, m.provenance]);

  async function handleUpload(file) {
    const busy = modal(`<h3 style="margin-top:0">Reading ${esc(file.name)}…</h3><p class="small">Nothing changes until you confirm.</p>`);
    let parsed;
    try { parsed = await window.BulkUpload.parseFile(file); }
    catch (err) {
      busy.remove();
      modal(`<h3 style="margin-top:0">Could not use this file</h3><p>${esc(err.message)}</p><div class="row-end"><button class="btn" data-close>OK</button></div>`);
      return;
    }
    busy.remove();
    const sm = parsed.summary, tableNames = { benchmarks: 'Competitor benchmarks', peerMonthly: 'Peer scenarios', facts: 'Company facts', sources: 'Source register', dictionary: 'Data dictionary', guide: 'Workbook guide', dashboardSummary2025: 'Dashboard summary' };
    const w = parsed.warnings;
    const bg = modal(`<h3 style="margin-top:0">Preview: ${esc(file.name)}</h3>
      <table>
        <tr><td>Monthly rows found</td><td><b>${sm.monthlyRows}</b>${sm.range ? ` (${mLong(sm.range[0])} – ${mLong(sm.range[1])})` : ''}</td></tr>
        ${sm.monthlyRows ? `<tr><td>New site-months</td><td><b>${sm.added}</b></td></tr><tr><td>Updated site-months</td><td><b>${sm.updated}</b></td></tr>
        <tr><td>Sites</td><td>${sm.sites.map(esc).join(', ')}${sm.newSites.length ? ` <span class="badge up">new: ${sm.newSites.map(esc).join(', ')}</span>` : ''}</td></tr>
        <tr><td>Cooling review flags</td><td>${sm.flags}</td></tr>` : ''}
        ${sm.tables.length ? `<tr><td>Other tables replaced</td><td>${sm.tables.map(t => tableNames[t] || t).join(', ')}</td></tr>` : ''}
      </table>
      ${w.length ? `<details class="warn-list"><summary>${w.length} warning(s)</summary><ul>${w.slice(0, 40).map(x => `<li>${esc(x)}</li>`).join('')}${w.length > 40 ? `<li>…and ${w.length - 40} more</li>` : ''}</ul></details>` : '<p class="small up-good">No problems found.</p>'}
      ${sm.monthlyRows ? `<label>How should the monthly rows be used?</label>
      <select id="up-mode"><option value="merge">Merge: update matching months, add new ones (recommended)</option><option value="replace">Replace: use only the uploaded monthly rows</option></select>` : ''}
      <p class="small">The upload is saved in this browser and can be removed again under Reports. Uploaded rows without a provenance column are labelled "UPLOADED".</p>
      <div class="row-end"><button class="btn ghost" data-close>Cancel</button><button class="btn" id="up-apply" ${sm.monthlyRows || sm.tables.length ? '' : 'disabled'}>Update dashboard</button></div>`);
    bg.querySelector('#up-apply').onclick = () => {
      const mode = bg.querySelector('#up-mode')?.value || 'merge';
      try { window.BulkUpload.apply(file.name, parsed, mode); }
      catch (err) { alert(err.message); return; }
      sessionStorage.setItem('bulk2.flash', `Dashboard updated from ${file.name}.`);
      location.hash = 'overview'; location.reload();
    };
  }

  function reportTable(months) {
    return `<table><tr><th>Site</th><th>IT (MWh)</th><th>Cooling (MWh)</th><th>Other (MWh)</th><th>Total (MWh)</th><th>PUE</th><th>Spot-index cost (NOK M)</th><th>Cooling flags</th></tr>
      ${SITES.map(s => { const a = agg(rows(months, [s])); return `<tr><td>${s}</td><td>${a.it.toFixed(0)}</td><td>${a.cool.toFixed(0)}</td><td>${a.other.toFixed(0)}</td><td>${a.total.toFixed(0)}</td><td>${a.pue.toFixed(3)}</td><td>${(a.cost / 1e6).toFixed(2)}</td><td>${flagged(months, [s]).length}</td></tr>`; }).join('')}
      ${(() => { const a = agg(rows(months)); return `<tr><th>All sites</th><th>${a.it.toFixed(0)}</th><th>${a.cool.toFixed(0)}</th><th>${a.other.toFixed(0)}</th><th>${a.total.toFixed(0)}</th><th>${a.pue.toFixed(3)}</th><th>${(a.cost / 1e6).toFixed(2)}</th><th>${flagged(months).length}</th></tr>`; })()}</table>`;
  }
  const csv = (name, header, data) => {
    const q = v => (/[",\n]/.test(String(v ?? '')) ? `"${String(v).replace(/"/g, '""')}"` : (v ?? ''));
    const blob = new Blob(['\ufeff' + [header, ...data].map(l => l.map(q).join(',')).join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  };
  function downloadReport(r) {
    if (r.type === 'XLSX') {
      csv(r.title.replace(/[^\w]+/g, '_') + '.csv', MONTHLY_HEADER, monthlyCsvRows());
      return;
    }
    let months, extra = '';
    if (r.kind === 'monthly') months = [r.month];
    else if (r.kind === 'ytd') months = MONTHS.filter(m => m.startsWith(LAST.slice(0, 4)));
    else if (r.kind === 'annual') months = MONTHS.filter(m => m.startsWith('2025'));
    else if (r.kind === 'anomaly') {
      months = [r.month];
      const x = rows([r.month], [r.site])[0];
      extra = `<h2>${r.site} – cooling review, ${mLong(r.month)}</h2><table><tr><th>Actual cooling</th><td>${x.coolingMWh.toFixed(0)} MWh</td></tr><tr><th>Expected cooling (model)</th><td>${x.expectedCoolingMWh.toFixed(0)} MWh</td></tr><tr><th>Residual</th><td>${x.residualMWh.toFixed(0)} MWh (${pctS(x.residualPct)})</td></tr><tr><th>IT load</th><td>${x.itLoadMW.toFixed(1)} MW</td></tr><tr><th>Outdoor temperature</th><td>${x.tempC} °C</td></tr><tr><th>Excess cost at spot index</th><td>NOK ${(excessNOK(x) / 1e6).toFixed(2)}M</td></tr></table>
        <h3>Suggested actions</h3><ul>${ACTIONS.map(a => `<li>${a}</li>`).join('')}</ul>`;
    } else months = MONTHS.slice(-3);
    const fl = flagged(months);
    if (r.kind !== 'anomaly' && fl.length) extra += `<h3>Cooling review flags</h3><ul>${fl.map(x => `<li>${x.site}, ${mLong(x.month)}: ${pctS(x.residualPct)} vs. expected</li>`).join('')}</ul>`;
    const w = window.open('', '_blank');
    if (!w) { alert('Please allow pop-ups to open the report.'); return; }
    w.document.write(`<html><head><title>${esc(r.title)}</title><style>body{font-family:Arial,sans-serif;padding:30px;color:#111}table{border-collapse:collapse;margin:10px 0}td,th{border:1px solid #ccc;padding:6px 10px;text-align:left;font-size:13px}.warn{background:#fff4d6;border:1px solid #f0c36d;padding:8px 12px;font-size:12px}</style></head><body>
      <h1>${esc(r.title)}</h1><p style="color:#666">Bulk Energy Intelligence · ${fmtDate(r.date)} · period ${rangeLabel(months)}</p>
      <p class="warn"><b>CONFIDENTIAL – internal Bulk use only.</b> Prototype report. All operational values are SYNTHETIC (source: ${esc(D.source)}). Cost is facility MWh × a synthetic price index, not Bulk's invoice.</p>
      ${reportTable(months)}${extra}<p style="color:#666;font-size:12px">Use "Save as PDF" in the print dialog.</p><script>setTimeout(()=>print(),300)<\/script></body></html>`);
    w.document.close();
  }
  function modal(html) {
    const bg = document.createElement('div'); bg.className = 'modal-bg'; bg.innerHTML = `<div class="modal">${html}</div>`;
    document.body.appendChild(bg);
    bg.addEventListener('click', e => { if (e.target === bg || e.target.closest('[data-close]')) bg.remove(); });
    return bg;
  }
  function newReportModal() {
    const bg = modal(`<h3 style="margin-top:0">New report</h3>
      <label>Report type</label><select id="rt"><option value="monthly">Monthly energy report (PDF)</option><option value="ytd">Year-to-date summary (PDF)</option><option value="annual">Annual energy report 2025 (PDF)</option><option value="anomaly">Cooling anomaly analysis (PDF)</option><option value="capacity">Monthly data export (XLSX/CSV)</option></select>
      <label>Month (monthly report)</label><select id="rm">${MONTHS.slice().reverse().map(m => `<option value="${m}">${mLong(m)}</option>`).join('')}</select>
      <label>Anomaly (anomaly report)</label><select id="ra">${flagged().map(r => `<option value="${r.site}|${r.month}">${r.site} – ${mLong(r.month)}</option>`).join('')}</select>
      <div class="row-end"><button class="btn ghost" data-close>Cancel</button><button class="btn" id="rc">Generate</button></div>`);
    bg.querySelector('#rc').onclick = () => {
      const kind = bg.querySelector('#rt').value, month = bg.querySelector('#rm').value, [aSite, aMonth] = bg.querySelector('#ra').value.split('|');
      const r = { monthly: { title: `Monthly energy report – ${mLong(month)}`, type: 'PDF', kind, month },
        ytd: { title: `${LAST.slice(0, 4)} year-to-date summary`, type: 'PDF', kind },
        annual: { title: 'Annual energy report 2025', type: 'PDF', kind },
        anomaly: { title: `${aSite} cooling anomaly analysis – ${mLong(aMonth)}`, type: 'PDF', kind, site: aSite, month: aMonth },
        capacity: { title: 'Monthly data export', type: 'XLSX', kind } }[kind];
      r.date = today;
      store.set('reports', [r, ...store.get('reports', [])]);
      bg.remove(); render();
    };
  }
  function detailsModal(month) {
    const s = state.site, x = rows([month], [s])[0];
    const prevRows = rows(MONTHS.slice(Math.max(0, MONTHS.indexOf(month) - 3), MONTHS.indexOf(month)), [s]);
    const avg = k => prevRows.reduce((t, r) => t + r[k], 0) / (prevRows.length || 1);
    modal(`<h3 style="margin-top:0">Unusual increase in cooling energy – ${s}, ${mLong(month)}</h3>
      <table><tr><th></th><th>${mShort(month)}</th><th>Prior 3-month avg.</th></tr>
        <tr><td>Cooling (MWh)</td><td>${x.coolingMWh.toFixed(0)}</td><td>${avg('coolingMWh').toFixed(0)}</td></tr>
        <tr><td>Expected cooling (MWh)</td><td>${x.expectedCoolingMWh.toFixed(0)}</td><td>${avg('expectedCoolingMWh').toFixed(0)}</td></tr>
        <tr><td>IT load (MW)</td><td>${x.itLoadMW.toFixed(1)}</td><td>${avg('itLoadMW').toFixed(1)}</td></tr>
        <tr><td>Outdoor temp. (°C)</td><td>${x.tempC.toFixed(1)}</td><td>${avg('tempC').toFixed(1)}</td></tr>
        <tr><td>PUE</td><td>${x.pue.toFixed(3)}</td><td>${avg('pue').toFixed(3)}</td></tr></table>
      <p><b>Residual:</b> ${x.residualMWh.toFixed(0)} MWh (${pctS(x.residualPct)}) above the model. The review threshold is +${THRESH * 100}%.</p>
      <p><b>Excess cost at the spot index:</b> NOK ${(excessNOK(x) / 1e6).toFixed(2)}M (${x.priceNOKMWh.toFixed(0)} NOK/MWh).</p>
      <p class="small">This is an AI-assisted flag on synthetic data. It points to where to look and does not diagnose a failure. Have site engineers validate it with real metering before acting.</p>
      <div class="row-end"><button class="btn" data-close>Close</button></div>`);
  }

  /* ---------- data & sources ---------- */
  pages.data = () => {
    const t = state.dataTab;
    let body = '';
    if (t === 'facts') body = `<table><thead><tr><th>Entity</th><th>Period</th><th>Metric</th><th>Value</th><th>Class</th><th>Caveat</th></tr></thead><tbody>${D.facts.map(f => `<tr><td>${esc(f.entity)}</td><td>${esc(f.period)}</td><td>${esc(f.metric)}</td><td><b>${esc(f.value ?? '')}</b> ${esc(f.unit)}</td><td>${esc(f.classification)}</td><td class="small">${esc(f.caveat)}</td></tr>`).join('')}</tbody></table>`;
    if (t === 'sources') body = `<table><thead><tr><th>ID</th><th>Organisation</th><th>Document</th><th>Type</th><th>Use and limitation</th></tr></thead><tbody>${D.sources.map(s => `<tr><td>${esc(s.id)}</td><td>${esc(s.organization)}</td><td>${/^https?:/.test(s.document || '') ? `<a href="${esc(s.document)}" target="_blank" rel="noopener">${esc(s.document)}</a>` : esc(s.document)}</td><td>${esc(s.type)}</td><td class="small">${esc(s.use)}</td></tr>`).join('')}</tbody></table>`;
    if (t === 'dictionary') body = `<table><thead><tr><th>Sheet</th><th>Field(s)</th><th>Unit</th><th>Method</th><th>Notes</th></tr></thead><tbody>${D.dictionary.map(d => `<tr><td>${esc(d.sheet)}</td><td>${esc(d.fields)}</td><td>${esc(d.unit)}</td><td>${esc(d.method)}</td><td class="small">${esc(d.notes)}</td></tr>`).join('')}</tbody></table>`;
    if (t === 'guide') body = D.guide.map(g => `<div class="rank" style="display:block"><b>${esc(g.step)}</b><div class="small" style="margin-top:4px;color:#c9ccda">${esc(g.text)}</div></div>`).join('');
    if (t === 'check') {
      const S = D.dashboardSummary2025, y = MONTHS.filter(m => m.startsWith('2025'));
      const lines = [['Annual facility energy (MWh)', a => a.total, 0], ['Annual IT energy (MWh)', a => a.it, 0], ['Annual cooling energy (MWh)', a => a.cool, 0], ['Annual weighted PUE', a => a.pue, 4], ['Spot-index electricity cost (NOK)', a => a.cost, 0]];
      body = `<p class="small">The dashboard recalculates the workbook's 2025 "Dashboard" sheet from the monthly rows. Matching values show the import is complete.</p><table><thead><tr><th>Metric</th>${SITES.map(s => `<th>${s} workbook</th><th>${s} dashboard</th>`).join('')}</tr></thead><tbody>
        ${lines.map(([l, f, d]) => `<tr><td>${l}</td>${SITES.map(s => { const w = S[l]?.[s], v = f(agg(rows(y, [s]))); const ok = Math.abs(w - v) / Math.abs(w) < 1e-4; return `<td>${(+w).toFixed(d)}</td><td class="${ok ? 'up-good' : 'up-bad'}">${v.toFixed(d)} ${ok ? '✓' : '✗'}</td>`; }).join('')}</tr>`).join('')}</tbody></table>`;
    }
    app.innerHTML = head('Data &amp; Sources', `Where every number comes from · ${esc(D.source)}`) +
      `<div class="banner"><span style="color:${C.amber}">${ic('warn', 22)}</span><div><b>Prototype on synthetic operating data</b><p>All monthly site measurements (energy, PUE, temperature, prices, cost) are simulated in the workbook. Verified items are N01's reported 59 MW IT capacity, Bulk's group financials and design PUE target, and the peers' published annual PUE. Real site-level data needs Bulk's permission.</p></div></div>` +
      tabs('dataTab', [['facts', 'Verified company facts'], ['sources', 'Source register'], ['dictionary', 'Data dictionary'], ['guide', 'Workbook guide'], ['check', 'Import check']]) +
      `<div class="card" data-info-key="data:${t}">${body}</div>`;
  };

  pages.settings = () => {
    app.innerHTML = head('Settings', 'Data and preferences') + `
      <div class="card" style="font-size:13px;line-height:1.7"><h3>Data file</h3>
        <p>Loaded from <code>js/data.js</code>, generated from <b>${esc(D.source)}</b> on ${esc(D.generated)}: ${D.monthly.length} site-months (${SITES.length} sites, ${mLong(MONTHS[0])} – ${mLong(LAST)}), ${D.benchmarks.length} competitor benchmarks, ${D.peerMonthly.length} peer scenario months, ${D.facts.length} company facts.</p>
        <p>To update after editing the Excel file, run <code>python3 tools/excel_to_data.py</code> in the <code>Final Project</code> folder.</p>
        <p>Uploaded data files: ${UPLOADS.length ? UPLOADS.map(u => `<b>${esc(u.name)}</b>`).join(', ') + ' (manage them on the Reports page)' : 'none'}.</p>
        <p>Saved in this browser only: the period you selected, ticked AI actions, reports you created and uploaded data.</p>
        <button class="btn ghost" data-action="reset">Reset saved actions and reports</button></div>
      <div class="card" style="font-size:13px;line-height:1.7;margin-top:12px"><h3>Contact &amp; help</h3>
        <p>A contact person and help line for questions about the platform and the numbers will be added here.</p>
        <p class="small"><b>This will be developed later.</b></p></div>`;
  };

  /* ---------- info buttons: plain-language explanations for every card and page ---------- */
  const T12 = `${(THRESH * 100).toFixed(0)}%`, TGT = PUE_TARGET.toFixed(2);
  const e = (title, what, basis, read) => ({ title, what, basis, read });
  const PAGE_INFO = {
    overview: e('Overview', 'A one-page summary of how all Bulk data centres are performing: efficiency (PUE), energy use, cooling, estimated cost and the latest AI flags.', 'All sites combined for the period chosen in the top-right menu. Operating figures come from the project workbook and are synthetic.', 'Start here, then open Sites, Energy & Cost or AI Insights for detail. Click ⓘ on any card for an explanation.'),
    sites: e('Sites', 'A closer look at one data centre at a time.', 'Monthly data for the chosen site and period.', 'Choose the data centre with the buttons under the title. The tabs show performance, the energy breakdown, cost and prices, and the AI analysis for that site.'),
    energy: e('Energy & Cost', 'How much electricity all sites use each month and what it would cost at market prices.', 'Monthly energy per site × the market price index for each site\'s price zone.', 'Use it to see whether cost changes come from using more energy or from price movements.'),
    insights: e('AI Insights', 'Automatic checks that flag months where cooling used more energy than expected.', `A transparent formula estimates "expected" cooling from IT load and outdoor temperature. A month is flagged when actual cooling is more than ${T12} above it.`, 'A flag tells you where to look, not what is wrong. Site engineers should confirm before acting.'),
    capacity: e('Capacity', 'How much of the available power capacity is used, and how much room is left for new customers.', `N01's reported IT capacity (${N01_CAPACITY ? N01_CAPACITY.value + ' MW' : 'n/a'}) and the average IT load per site in the latest month.`, 'Headroom shrinking towards zero means expansion has to be planned. Capacity for OS-IX and DK01 is not verified, so only their load is shown.'),
    benchmarking: e('Benchmarking', 'How Bulk compares with other Nordic data centre operators.', 'Published annual PUE figures from competitors\' sustainability reports, plus synthetic monthly scenarios and electricity prices.', 'Click a competitor to compare it side by side with Bulk. The figures use different years and methods, so treat them as context, not a ranking.'),
    reports: e('Reports', 'Download ready-made reports, or upload a new data file to update the numbers.', 'Reports are generated from the data currently loaded in the dashboard.', 'PDF reports open in a print view ("Save as PDF"). Uploads show a preview before anything changes.'),
    data: e('Data & Sources', 'Where every number comes from, and which numbers are verified and which are simulated.', 'All sheets of the project workbook.', 'Use this page to check a source or a definition, or to show that the import matches the workbook.'),
    settings: e('Settings', 'Information about the loaded data, saved preferences, and help.', 'Settings are stored in this browser only.', ''),
  };
  const INFO = [
    // Overview
    [/^overview\|about this platform/, e('About this platform', 'A short introduction to what the dashboard is for, plus the key numbers for the selected period.', 'The same data as the cards below.', 'The bullet points update when you change the period.')],
    [/^overview\|average pue/, e('Average PUE (Power Usage Effectiveness)', 'How efficiently the data centres use electricity. PUE = total energy ÷ energy used by the IT equipment (servers, storage, network). 1.00 would mean every kWh goes to IT. 1.20 means 0.20 kWh extra for cooling, power losses and lighting for every kWh of IT.', 'Total facility MWh ÷ IT MWh for all sites in the selected period, weighted by energy.', `Lower is better. Bulk's design target for new builds is max. ${TGT}. The arrow compares with the previous comparable period: green = improvement, red = worse.`)],
    [/^overview\|energy consumption by site/, e('Energy consumption by site', 'How much electricity each data centre used in the selected period, split into what it was used for.', 'Monthly energy per site, summed over the period, in GWh (1 GWh = 1,000 MWh).', 'Purple = IT (customer servers), blue = cooling, grey = other overhead such as power losses and lighting. A taller bar means more energy. A larger blue and grey part relative to the purple means a higher (worse) PUE.')],
    [/^overview\|ai insights/, e('AI insights', `Automatic flags for months where cooling energy was more than ${T12} above what the model expected.`, 'Expected cooling = IT energy × (0.115 + 0.0055 × degrees above 8 °C + site adjustment). This is a simple, transparent formula, not a trained machine-learning model.', 'Click a flag to open AI Insights. "Potential saving" adds up the excess cooling in the flagged months. A flag points to where to look; engineers must confirm the cause.')],
    // KPIs shared across pages
    [/\|pue$/, e('PUE (Power Usage Effectiveness)', 'Total energy ÷ IT energy for this site. It shows how much extra energy the building needs on top of the servers themselves.', 'Total facility MWh ÷ IT MWh for the selected period, weighted by energy.', `Lower is better. The target for new builds is max. ${TGT}. Green arrow = improved vs. the previous period.`)],
    [/\|total energy$/, e('Total energy', 'All electricity used: IT + cooling + other overhead.', 'Sum of "Total facility (MWh)" for the selected period and sites, shown in GWh (1 GWh = 1,000 MWh).', 'More energy is not automatically bad. It usually means customers use more IT. Look at PUE to judge efficiency.')],
    [/\|it energy$/, e('IT energy', 'Electricity used by customers\' IT equipment. This is the "useful" energy the data centre exists to deliver.', 'Average IT load (MW) × hours in each month, summed for the period.', 'Grows when customers add servers. Rising IT energy usually means more business.')],
    [/\|cooling energy$/, e('Cooling energy', 'Electricity used by chillers, fans and pumps to keep the IT equipment cool. Usually the largest overhead.', 'Monthly cooling MWh for the selected period and sites.', `Compare it with total energy (its share). Months where cooling is more than ${T12} above the model's expectation are flagged in AI Insights.`)],
    [/\|cooling share$/, e('Cooling share', 'The part of all energy that goes to cooling.', 'Cooling MWh ÷ total facility MWh for the selected period.', 'Lower is better. The share is naturally higher in warm months. The arrow compares with the previous period (green = lower share).')],
    [/\|(estimated energy cost|total cost|estimated energy cost)$/, e('Estimated energy cost', 'What the electricity would cost at market prices.', 'Total MWh × the monthly market price index (NOK/MWh) in each site\'s price zone (NO1, NO2, DK1). The price index in the workbook is synthetic.', 'This is not Bulk\'s actual bill: grid tariffs, taxes, hedging and customer pass-through are not included. Use it to see how prices and consumption drive cost.')],
    [/\|avg\. price$/, e('Average electricity price', 'The average price paid per unit of electricity.', 'Estimated cost ÷ total energy, shown in NOK per kWh (1 NOK/kWh = 1,000 NOK/MWh).', 'Driven by the market. Winter months are usually more expensive. Red arrow = higher price than the previous period.')],
    [/\|price zone$/, e('Price zone', 'The Nordic power market is split into price zones, and each site buys electricity at its zone\'s price.', 'NO1 = Oslo area (OS-IX), NO2 = southern Norway (N01), DK1 = western Denmark (DK01).', 'Prices can differ a lot between zones, so the same consumption costs different amounts at different sites.')],
    [/\|avg\. outdoor temp/, e('Average outdoor temperature', 'The average monthly outside temperature at the site.', 'Synthetic weather values from the workbook, not measured readings.', 'Warmer weather needs more cooling. The expected-cooling model adds cooling for every degree above 8 °C.')],
    // Sites
    [/^sites\|pue trend/, e('PUE trend', 'How the site\'s efficiency changed month by month.', 'Total ÷ IT energy for each month.', `The dashed line is Bulk's design target (max. ${TGT}). Points above the line are less efficient than the target. A rising line means efficiency is getting worse, which is often seasonal (warmer weather).`)],
    [/^sites\|energy breakdown/, e('Energy breakdown', 'What the site\'s electricity was used for in the latest month of the period.', 'IT, cooling and other MWh for that month. The centre shows the total in GWh.', 'A bigger purple (IT) share means a more efficient site, because less energy goes to overhead.')],
    [/^sites\|monthly energy breakdown/, e('Monthly energy breakdown', 'IT, cooling and other energy for each month.', 'Monthly MWh for the selected site, shown in GWh.', 'Watch the blue (cooling) part. If it grows while purple (IT) stays flat, efficiency is falling.')],
    [/^sites\|monthly cost/, e('Monthly cost and market price', 'How the estimated cost and the electricity price moved each month.', 'Bars = estimated cost in NOK million (left axis). Line = market price in NOK/MWh (right axis).', 'If cost follows the price line closely, changes come from the market rather than from using more energy.')],
    [/^sites\|cooling residual/, e('Cooling residual vs. model', 'How much the actual cooling differed from what the model expected, in percent, for each month.', `(Actual − expected cooling) ÷ expected cooling. Expected cooling depends on IT load and outdoor temperature.`, `Amber bars above the dashed ${T12} line are flagged for review. Negative bars mean less cooling than expected.`)],
    [/^sites\|ai analysis for/, e('AI analysis for this site', 'The cooling flags for this site and how its PUE compares with the design target.', `Cooling residual > ${T12} = review flag. PUE target max. ${TGT}.`, 'Click a flag to see it in AI Insights.')],
    [/^sites\|no data for/, e('No data', 'The chosen site has no rows in the selected period.', 'Data loaded in the dashboard, including any uploads.', 'Choose another period or data centre.')],
    // Energy & Cost
    [/^energy\|monthly energy/, e('Monthly energy and cost', 'Total electricity use and estimated cost for all sites, month by month.', 'Bars = energy in GWh (left axis). Line = estimated cost in NOK million (right axis).', 'If the line rises while the bars stay level, prices went up. If both rise, the sites used more energy.')],
    [/^energy\|cost by site/, e('Cost by site', 'Energy, efficiency, price and estimated cost for each data centre in the selected period.', 'Monthly site data summed for the period. Average price = cost ÷ energy.', 'Differences in price zone can matter as much as differences in consumption.')],
    // AI Insights
    [/^insights\|unusual increase in cooling/, e('Cooling alert', `The most recent month at this site where cooling was more than ${T12} above the model's expectation.`, 'Actual cooling vs. expected cooling (from IT load and temperature).', '"View details" shows the numbers behind the flag compared with the three months before.')],
    [/^insights\|no unusual cooling/, e('No cooling alerts', `Every month at this site is within ±${T12} of expected cooling.`, 'Actual vs. expected cooling for all months.', 'No action needed. Flags at other sites are listed further down.')],
    [/^insights\|actual vs\. expected cooling/, e('Actual vs. expected cooling', 'The site\'s cooling energy each month compared with what the model expected.', 'Purple line = actual cooling (synthetic data). Dashed line = expected cooling given IT load and outdoor temperature.', 'Where the purple line rises clearly above the dashed line, cooling used more than it should. Purple-shaded months are flagged (+%). The light shading marks the selected period.')],
    [/^insights\|suggested actions/, e('Suggested actions', 'A follow-up checklist for the flag.', 'Standard first checks for unexpected cooling consumption.', 'Tick actions as they are done. Ticks are saved in this browser only.')],
    [/^insights\|estimated impact/, e('Estimated impact', 'What fixing the cause could save.', 'Energy = excess cooling (actual − expected) in the flagged months. Cost = excess MWh × that month\'s price. Share = excess ÷ cooling energy in those months.', 'An upper estimate: it assumes all the excess can be removed.')],
    [/^insights\|flags at other sites/, e('Flags at other sites', 'Cooling flags at the other data centres.', `Same rule: cooling > ${T12} above expected.`, 'Click one to switch to that site.')],
    // Capacity
    [/^capacity\|installed capacity/, e('Installed capacity (N01)', 'The maximum IT power the N01 site is built to supply to customers.', `Reported operating IT capacity for 2025 (Menon Economics, Aug 2026). The ${N01_PLANNED ? N01_PLANNED.value + ' MW' : ''} planned by 2031 is a projection, not built.`, 'It is capacity, not usage. Compare it with "Utilised" to see how full the site is.')],
    [/^capacity\|utilised/, e('Utilised', 'How much IT power customers are using.', 'Sum of the average IT load (MW) across all sites in the latest month (synthetic).', 'This is an average. Short peaks can be higher.')],
    [/^capacity\|available headroom/, e('Available headroom (N01)', 'How much more customer load N01 could take before it is full.', 'N01 reported capacity − N01 average IT load in the latest month.', 'Below about 25% headroom, expansion should be planned. The site card shows the status.')],
    [/^capacity\|average it load by site/, e('Average IT load by site', 'How customer IT load developed over time at each site.', 'Average IT load in MW per month. The dashed line is N01\'s reported capacity.', 'A line approaching the dashed line means the site is filling up.')],
    [/^capacity\|.+ · /, e('Site capacity', 'How full this data centre is.', 'Average IT load in the latest month. For N01 this is compared with its reported capacity.', 'Bar and status: below 75% = healthy headroom, 75–85% = plan expansion, above 85% = critical. Capacity for OS-IX and DK01 is not verified, so only their load is shown.')],
    // Benchmarking
    [/^benchmarking\|reported annual pue/, e('Reported annual PUE', 'The efficiency figures competitors publish in their sustainability reports, next to Bulk.', 'Annual PUE from each operator\'s own report. Bulk\'s value is the 2025 figure from the synthetic data (green). Verne\'s Iceland and Finland sites are combined.', `Lower is better. The dashed line is Bulk's design target ${TGT}. Click a bar to compare. Years, scopes and methods differ, so small gaps are not meaningful.`)],
    [/^benchmarking\|monthly pue scenarios/, e('Monthly PUE scenarios', 'Illustrative month-by-month PUE lines for Bulk and the competitors.', 'Competitor lines are synthetic: generated around each competitor\'s published annual PUE. They are not real monthly data.', 'Use it to see seasonal patterns. Click a line or a legend name to compare.')],
    [/^benchmarking\|market price index by zone/, e('Market price by zone', 'The monthly electricity price in each site\'s price zone.', 'Synthetic price index in NOK/MWh from the workbook.', 'Shows how much more expensive one zone can be than another, and when prices peak.')],
    [/^benchmarking\|average pue/, e('Average PUE list', 'Each competitor\'s PUE next to Bulk.', 'Published annual figures (or synthetic scenario averages on that tab). The peer average leaves out EcoDataCenter Falun to avoid counting the same operator twice. Verne = simple average of Iceland and Finland.', 'Click a row (+) to add it to the side-by-side comparison, up to 4.')],
    [/^benchmarking\|average price/, e('Average price', 'The average market price in each zone over all months.', 'Mean of the monthly price index per zone.', 'A higher average price makes the same efficiency more valuable.')],
    [/^benchmarking\|(side-by-side comparison|click a competitor)/, e('Side-by-side comparison', 'A detailed comparison of Bulk with the competitors you selected.', 'Published figures plus the workbook\'s notes on scope, status and cautions. "If Bulk ran at this PUE" = Bulk\'s 2025 IT energy × (competitor PUE − Bulk PUE).', 'Green difference = Bulk is more efficient. Read the caution row before drawing conclusions.')],
    [/^benchmarking\|electricity price n01/, e('Electricity price N01', 'The monthly market price in N01\'s price zone (NO2).', 'Synthetic price index in NOK/MWh.', 'Peaks show when energy efficiency saves the most money.')],
    [/^benchmarking\|key takeaway/, e('Key takeaway', 'A short written conclusion of the comparison.', 'Generated automatically from the numbers on this page, including any competitors you selected.', 'Always read it with the comparability notes in mind.')],
    [/^benchmarking\|comparability notes/, e('Comparability notes', 'Why the competitor figures are not directly comparable.', 'Scope, status and cautions recorded in the workbook for each source.', 'Check this before using a figure in a presentation.')],
    // Reports
    [/^reports\|upload a report/, e('Upload data', 'Update the dashboard with a new Excel (.xlsx) or CSV file.', 'The file is read in your browser. Missing values are calculated with the workbook\'s formulas.', 'You get a preview first. "Merge" updates matching months and adds new ones. Uploads are saved only in this browser and can be removed again.')],
    [/^reports\|available reports/, e('Available reports', 'Ready-made reports generated from the current data.', 'The data loaded in the dashboard, including uploads.', 'PDF reports open in a print view. Choose "Save as PDF". The data export downloads as CSV for Excel.')],
    // Data & Sources
    [/^data\|prototype on synthetic/, e('Data status', 'Which numbers are verified and which are simulated.', 'The workbook\'s own classification.', 'Monthly operating figures are synthetic. Company facts and competitor PUE are taken from published reports.')],
    [/^data\|data:facts/, e('Verified company facts', 'Facts about Bulk from annual reports and studies, such as revenue, capacity and targets.', 'Workbook sheet "Real Company Facts".', 'The class column shows whether a figure is reported, projected, a target or an estimate.')],
    [/^data\|data:sources/, e('Source register', 'Every source used, with how it may and may not be used.', 'Workbook sheet "Sources & Methods".', 'Rows starting with M describe formulas and model assumptions.')],
    [/^data\|data:dictionary/, e('Data dictionary', 'The definition and unit of each data field.', 'Workbook sheet "Data Dictionary".', 'Use it when you are unsure what a column means.')],
    [/^data\|data:guide/, e('Workbook guide', 'The workbook\'s own instructions and limitations.', 'Workbook sheet "START HERE".', '')],
    [/^data\|data:check/, e('Import check', 'Proof that the dashboard reads the workbook correctly.', 'The dashboard recalculates the workbook\'s 2025 summary from the monthly rows.', 'A ✓ means the dashboard value matches the workbook.')],
    // Settings
    [/^settings\|data file/, e('Data file', 'Which data is loaded and what is stored in this browser.', '', '')],
    [/^settings\|contact/, e('Contact & help', 'Where to get help with the platform.', '', 'This will be developed later.')],
  ];
  let currentPage = 'overview';
  const findInfo = (page, key) => INFO.findIndex(([re]) => re.test(`${page}|${key.toLowerCase().replace(/\s+/g, ' ').trim()}`));
  const infoBtn = (data, cls = '') => `<button type="button" class="info-btn ${cls}" data-info="${data}" aria-label="What does this show?">${ic('info', 14)}</button>`;
  function decorateInfo(page) {
    app.querySelectorAll('.card, .banner').forEach(card => {
      if (card.querySelector(':scope > .info-btn')) return;
      const key = card.dataset.infoKey || (card.querySelector('.label, h3, b')?.textContent || '').trim();
      const id = findInfo(page, key);
      if (id < 0) console.warn('No info text for', page, key);
      card.classList.add('has-info');
      card.insertAdjacentHTML('afterbegin', infoBtn(id < 0 ? `page:${page}` : id, 'corner'));
    });
  }
  const pop = document.createElement('div');
  pop.className = 'info-pop'; pop.setAttribute('role', 'tooltip'); pop.hidden = true;
  document.body.appendChild(pop);
  let pinned = null;
  function showInfo(btn) {
    const id = btn.dataset.info, en = String(id).startsWith('page:') ? PAGE_INFO[id.slice(5)] : INFO[+id]?.[1];
    if (!en) return;
    const sec = (h, t) => t ? `<div class="sec"><b>${h}</b>${t}</div>` : '';
    pop.innerHTML = `<h4>${esc(en.title)}</h4>${sec('What it means', esc(en.what))}${sec('Based on', esc(en.basis))}${sec('How to read it', esc(en.read))}`;
    pop.hidden = false;
    const r = btn.getBoundingClientRect(), w = pop.offsetWidth, h = pop.offsetHeight;
    let left = Math.min(Math.max(8, r.right - w), innerWidth - w - 8), top = r.bottom + 6;
    if (top + h > innerHeight - 8) top = Math.max(8, r.top - h - 6);
    pop.style.left = left + 'px'; pop.style.top = top + 'px';
  }
  const hideInfo = () => { pop.hidden = true; pinned = null; };
  document.addEventListener('mouseover', ev => { const b = ev.target.closest('.info-btn'); if (b && !pinned) showInfo(b); });
  document.addEventListener('mouseout', ev => { const b = ev.target.closest('.info-btn'); if (b && !pinned && !b.contains(ev.relatedTarget)) pop.hidden = true; });
  document.addEventListener('focusin', ev => { const b = ev.target.closest('.info-btn'); if (b) showInfo(b); });
  document.addEventListener('click', ev => {
    const b = ev.target.closest('.info-btn');
    if (b) { ev.preventDefault(); ev.stopPropagation(); if (pinned === b) hideInfo(); else { pinned = b; showInfo(b); } return; }
    if (!pop.contains(ev.target)) hideInfo();
  }, true);
  document.addEventListener('keydown', ev => { if (ev.key === 'Escape') hideInfo(); if ((ev.key === 'Enter' || ev.key === ' ') && ev.target.classList?.contains('nav-info')) { ev.preventDefault(); showInfo(ev.target); pinned = ev.target; } });
  window.addEventListener('scroll', () => { if (!pop.hidden) hideInfo(); }, true);
  document.querySelectorAll('.sidebar a[data-page]').forEach(a => a.insertAdjacentHTML('beforeend', `<span class="info-btn nav-info" role="button" tabindex="0" aria-label="About this page" data-info="page:${a.dataset.page}">${ic('info', 13)}</span>`));

  /* ---------- data centre picker (Sites, AI Insights) ---------- */
  const sitePicker = () => `<div class="site-picker" role="group" aria-label="Choose data centre"><span class="sp-label">${ic('sites', 14)} Data centre</span>${SITES.map(s => `<button type="button" class="sp ${s === state.site ? 'on' : ''}" data-pick-site="${esc(s)}" aria-pressed="${s === state.site}"><b>${esc(s)}</b><span>${esc(SITE_INFO[s]?.location || '')}</span></button>`).join('')}</div>`;

  /* ---------- AI Copilot: rule-based decision support computed from the loaded data (no external AI) ---------- */
  ICONS.bot = '<rect x="4" y="8" width="16" height="12" rx="3"/><path d="M12 4v4M9 13h.01M15 13h.01M9.5 16.5h5"/><circle cx="12" cy="3.5" r="1"/>';
  ICONS.send = '<path d="M22 2L11 13M22 2l-7 20-4-9-9-4z"/>';
  ICONS.expand = '<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>';
  ICONS.shrink = '<path d="M4 14h6v6M20 10h-6V4M14 10l7-7M3 21l7-7"/>';
  ICONS.x = '<path d="M18 6L6 18M6 6l12 12"/>';
  const isMac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
  const SUGGESTIONS = ['Why did cooling energy increase?', 'Which site has the highest optimisation potential?', 'What capacity risks exist?', 'How do we compare with Nordic peers?'];
  const MORE = ['What drives our energy cost?', 'How efficient is N01?', 'What can you help with?'];
  const SITE_WORDS = { oslo: 'OS-IX', kristiansand: 'N01', vennesla: 'N01', esbjerg: 'DK01' };
  const ctxSite = () => (['sites', 'insights'].includes(currentPage) ? state.site : null);
  const ctxLabel = () => `Context: ${ctxSite() ? `${ctxSite()} (${SITE_INFO[ctxSite()]?.location || ''})` : 'All sites'} · ${PERIODS[state.period] || ''}`;

  function scope(q) {
    const t = q.toLowerCase();
    let site = SITES.find(s => new RegExp(`\\b${s.toLowerCase().replace('-', '[- ]?')}\\b`).test(t)) || Object.entries(SITE_WORDS).find(([w]) => t.includes(w))?.[1] || null;
    if (!site && /\b(all|every|each|which) site|portfolio|bulk\b/.test(t)) site = null; else if (!site) site = ctxSite();
    let months = periodMonths().cur, plabel = PERIODS[state.period];
    if (/\b2025\b/.test(t)) { months = MONTHS.filter(m => m.startsWith('2025')); plabel = 'Full year 2025'; }
    else if (/\b2026\b/.test(t)) { months = MONTHS.filter(m => m.startsWith('2026')); plabel = 'Year 2026'; }
    else if (/last month|latest month|this month/.test(t)) { months = [LAST]; plabel = mLong(LAST); }
    else if (/all (data|months|time)|ever|history/.test(t)) { months = MONTHS.slice(); plabel = 'All data'; }
    return { site, sites: site ? [site] : SITES, months, plabel };
  }
  const avgOf = (rs, k) => rs.length ? rs.reduce((t, r) => t + r[k], 0) / rs.length : 0;
  const prior3 = r => rows(MONTHS.slice(Math.max(0, MONTHS.indexOf(r.month) - 3), MONTHS.indexOf(r.month)), [r.site]);
  const answer = (o) => o; // {title, obs:[], expl:[], next:[], links:[[label,page,site]], basis}

  const INTENTS = [
    { re: /cool|anomal|flag|alert|residual|unusual|spike/, fn: s => {
      let fl = flagged(s.months, s.sites), outside = false;
      if (!fl.length) { fl = flagged(MONTHS, s.sites); outside = fl.length > 0; }
      if (!fl.length) return answer({ title: `No unusual cooling ${s.site ? `at ${s.site}` : 'at any site'}`, obs: [`Every month ${s.site ? `at ${s.site}` : ''} is within ±${(THRESH * 100).toFixed(0)}% of expected cooling.`], expl: [], next: ['No action needed. Cooling share is shown on Energy & Cost.'], basis: 'Actual vs. expected cooling, all months.' });
      const obs = [], expl = new Set();
      if (outside) obs.push(`No flags in ${s.plabel}; showing the most recent flags outside it.`);
      fl.slice(0, 3).forEach(r => {
        const p = prior3(r), it0 = avgOf(p, 'itLoadMW'), t0 = avgOf(p, 'tempC'), c0 = avgOf(p, 'coolingMWh');
        obs.push(`<b>${r.site}, ${mLong(r.month)}:</b> cooling ${r.coolingMWh.toFixed(0)} MWh vs. ${r.expectedCoolingMWh.toFixed(0)} expected (<b>${pctS(r.residualPct, 0)}</b>, ${Math.max(r.residualMWh, 0).toFixed(0)} MWh excess, about NOK ${(excessNOK(r) / 1e6).toFixed(2)}M).` +
          (p.length ? ` Compared with the 3 months before: cooling ${pctS((r.coolingMWh - c0) / c0, 0)}, IT load ${pctS((r.itLoadMW - it0) / it0, 1)}, outdoor temperature ${r.tempC.toFixed(1)} °C vs. ${t0.toFixed(1)} °C.` : ''));
        if (r.tempC - t0 > 3) expl.add('Warmer weather raised cooling demand. The model already adjusts for monthly average temperature, but short heat peaks inside the month are not captured.');
        if (p.length && (r.itLoadMW - it0) / it0 > 0.03) expl.add('Higher IT load, or denser racks creating hot spots. The model scales with IT energy but not with rack density.');
      });
      ['Cooling plant performance: setpoint drift, a pump or fan running harder than needed, or condenser/heat-exchanger fouling.', 'A maintenance event or configuration change, for example equipment in bypass or changed control settings.', 'A metering or data error. Check the meter readings before acting.'].forEach(x => expl.add(x));
      return answer({ title: `Why cooling energy increased${s.site ? ` at ${s.site}` : ''}`, obs, expl: [...expl], next: ['Check cooling setpoints, pumps and condenser for the flagged month(s).', 'Review the maintenance log for the same period.', 'Compare with the other sites in the same month.', 'Validate against permissioned metering before deciding.'],
        links: [['Open AI Insights', 'insights', fl[0].site]], basis: `Cooling residual = (actual − expected) ÷ expected; expected from IT energy and temperature (workbook M05/M06). Flag above +${(THRESH * 100).toFixed(0)}%. The anomalies were injected into synthetic data.` });
    } },
    { re: /optimi|potential|saving|improve|where.*(save|focus)|priorit|efficien.*(gain|opportun)/, fn: s => {
      const list = SITES.map(site => {
        const a = agg(rows(s.months, [site])); if (!a.it) return null;
        const gap = Math.max(0, a.pue - PUE_TARGET) * a.it, fl = flagged(s.months, [site]);
        const ex = fl.reduce((t, r) => t + Math.max(r.residualMWh, 0), 0);
        return { site, a, gap, gapNOK: gap * a.price, ex, exNOK: fl.reduce((t, r) => t + excessNOK(r), 0), n: fl.length };
      }).filter(Boolean).map(x => ({ ...x, pot: Math.max(x.gap, x.ex) })).sort((x, y) => y.pot - x.pot);
      if (!list.length) return answer({ title: 'No data', obs: [`No data for ${s.plabel}.`], expl: [], next: [] });
      const top = list[0];
      return answer({ title: `${top.site} has the highest optimisation potential`, obs: list.map((x, i) => `${i + 1}. <b>${x.site}</b>: PUE ${x.a.pue.toFixed(3)} (${x.a.pue > PUE_TARGET ? `${(x.a.pue - PUE_TARGET).toFixed(3)} above` : 'at or below'} the ${PUE_TARGET.toFixed(2)} target) → ${(x.gap / 1000).toFixed(1)} GWh about ${nokM(x.gapNOK)} to reach target.${x.n ? ` Plus ${x.n} cooling flag(s) with ${(x.ex / 1000).toFixed(2)} GWh excess (about ${nokM(x.exNOK)}).` : ''}`),
        expl: [top.ex > top.gap ? `${top.site} ranks first because of the excess cooling in its flagged month(s) (${(top.ex / 1000).toFixed(2)} GWh), which is larger than any site's gap to the PUE target.` : top.a.it >= Math.max(...list.map(x => x.a.it)) ? `${top.site} ranks first mainly because of its size: ${(top.a.it / 1000).toFixed(0)} GWh of IT energy, so even a small PUE gap means a lot of energy.` : `${top.site} ranks first because it has the largest gap to the PUE target relative to its size.`, 'The two figures overlap (fixing excess cooling also lowers PUE), so the ranking uses the larger of the two for each site rather than adding them.'],
        next: [`Start with the flagged cooling months at ${top.site}; they are concrete and quick to check.`, 'Then review seasonal PUE (summer peaks) on the Sites page.'],
        links: [['Open Sites', 'sites', top.site]], basis: `${s.plabel}. Gap energy = (PUE − ${PUE_TARGET.toFixed(2)}) × IT energy; cost at each site's average price. Synthetic operating data.` });
    } },
    { re: /capacit|headroom|full|expan|space|room|utili[sz]|load growth/, fn: () => {
      const cap = N01_CAPACITY?.value, n = rows(MONTHS, ['N01']), last = n.at(-1), obs = [], expl = [], next = [];
      if (cap && last) {
        const y = n.slice(-12).map(r => r.itLoadMW), xs = y.map((_, i) => i), mx = avgOf(xs.map(v => ({ v })), 'v'), my = y.reduce((t, v) => t + v, 0) / y.length;
        const slope = xs.reduce((t, x, i) => t + (x - mx) * (y[i] - my), 0) / xs.reduce((t, x) => t + (x - mx) ** 2, 0);
        const u = last.itLoadMW / cap, to85 = slope > 0 ? (0.85 * cap - last.itLoadMW) / slope : Infinity, to100 = slope > 0 ? (cap - last.itLoadMW) / slope : Infinity;
        obs.push(`<b>N01</b> uses ${last.itLoadMW.toFixed(1)} of ${cap} MW reported IT capacity (<b>${(u * 100).toFixed(0)}%</b>) in ${mLong(LAST)}, leaving ${(cap - last.itLoadMW).toFixed(1)} MW headroom.`);
        obs.push(`IT load grew about ${(slope * 12).toFixed(1)} MW per year over the last 12 months.${isFinite(to85) ? ` At that pace N01 reaches 85% in about ${Math.max(0, Math.round(to85))} months and 100% in about ${Math.round(to100)} months.` : ' Load is not growing, so there is no near-term limit.'}`);
        expl.push('This is a straight-line trend on synthetic load. New customer contracts arrive in steps, so real timing can be much faster or slower.');
        if (N01_PLANNED) expl.push(`The planned ${N01_PLANNED.value} MW by 2031 is a projection, not built or contracted capacity.`);
        next.push(isFinite(to85) && to85 < 24 ? 'Start expansion planning for N01 now: the 85% level is within two years.' : 'Monitor N01 quarterly; headroom is adequate for now.');
      }
      ['OS-IX', 'DK01'].filter(x => SITES.includes(x)).forEach(x => { const r = rows([LAST], [x])[0]; if (r) obs.push(`<b>${x}</b>: ${r.itLoadMW.toFixed(1)} MW average IT load. Installed capacity is not verified, so the risk cannot be assessed.`); });
      next.push('Obtain verified capacity figures for OS-IX and DK01.');
      return answer({ title: 'Capacity risks', obs, expl, next, links: [['Open Capacity', 'capacity']], basis: 'N01 capacity: Menon Economics (Aug 2026), 2025. Load: synthetic monthly IT load.' });
    } },
    { re: /peer|compet|nordic|benchmark|compare|atnorth|verne|green mountain|ecodata|stack|rival|market/, fn: () => {
      const b = agg(rows(MONTHS.filter(m => m.startsWith('2025')))), comps = competitors().filter(x => typeof x.pue === 'number').sort((x, y) => x.pue - y.pue);
      const indep = comps.filter(x => !/Falun/.test(x.facility)), avg = indep.reduce((t, x) => t + x.pue, 0) / indep.length;
      const better = comps.filter(x => x.pue < b.pue);
      return answer({ title: 'How Bulk compares with Nordic peers', obs: [`Bulk's 2025 PUE is <b>${b.pue.toFixed(3)}</b>, ${((avg - b.pue) / avg * 100).toFixed(1)}% lower than the peer average of ${avg.toFixed(2)}.`,
        ...comps.map(x => `${esc(x.operator)} (${esc(x.facility)}, ${x.year}): ${x.pue.toFixed(x.merged ? 3 : 2)}, ${x.pue > b.pue ? `${(x.pue - b.pue).toFixed(3)} higher` : x.pue < b.pue ? `${(b.pue - x.pue).toFixed(3)} lower` : 'equal'} → ${(b.it * (x.pue - b.pue) / 1000).toFixed(1)} GWh/year difference at Bulk's scale.`),
        better.length ? `Lower than Bulk: ${better.map(x => esc(x.operator)).join(', ')}.` : 'No peer reports a lower PUE than Bulk\'s figure.'],
        expl: ['Bulk\'s figure is synthetic; the peers\' figures are published but from different years (2024/2025), scopes (portfolio vs. campus, some include non-electric energy) and climates (Iceland vs. Finland vs. Norway).', 'Differences under about 0.03 are within these method differences and should not be treated as real.'],
        next: ['Use the side-by-side comparison on Benchmarking and read the caution row.', 'Replace Bulk\'s synthetic PUE with measured data before external use.'],
        links: [['Open Benchmarking', 'benchmarking']], basis: 'Competitor Benchmarks sheet (verified annual PUE). STACK has no published figure and is excluded.' });
    } },
    { re: /cost|price|spend|nok|expens|bill|tariff|money/, fn: s => {
      const a = agg(rows(s.months, s.sites)), p = periodMonths().prev, pa = p && s.months === periodMonths().cur ? agg(rows(p, s.sites)) : null;
      const ms = s.months.map(m => ({ m, ...agg(rows([m], s.sites)) })).filter(x => x.total), hi = ms.reduce((x, y) => y.cost > x.cost ? y : x, ms[0] || {}), cheap = ms.reduce((x, y) => y.price < x.price ? y : x, ms[0] || {});
      const obs = [`Estimated cost ${s.site ? `at ${s.site}` : 'for all sites'}: <b>${nokM(a.cost)}</b> for ${(a.total / 1000).toFixed(1)} GWh in ${s.plabel}, an average of ${(a.price / 1000).toFixed(2)} NOK/kWh.`];
      if (hi.m) obs.push(`Highest month: ${mLong(hi.m)} (${nokM(hi.cost)}, ${hi.price.toFixed(0)} NOK/MWh). Cheapest power: ${mLong(cheap.m)} (${cheap.price.toFixed(0)} NOK/MWh).`);
      const expl = [];
      if (pa) {
        const vol = (a.total - pa.total) * pa.price, pr = (a.price - pa.price) * a.total;
        obs.push(`Vs. ${periodMonths().prevLabel.replace('vs. ', '')}: cost ${pctS((a.cost - pa.cost) / pa.cost)}; ${nokM(vol)} from using more energy and ${nokM(pr)} from price changes.`);
        expl.push(Math.abs(pr) > Math.abs(vol) ? 'The change is driven mostly by market prices, which Bulk cannot control directly.' : 'The change is driven mostly by higher consumption, i.e. customer IT growth.');
      }
      if (!s.site) SITES.forEach(x => { const b = agg(rows(s.months, [x])); if (b.total) obs.push(`${x}: ${nokM(b.cost)} (${(b.cost / a.cost * 100).toFixed(0)}%), ${(b.price / 1000).toFixed(2)} NOK/kWh in ${esc(zone(x))}.`); });
      expl.push('Prices differ between price zones (NO1, NO2, DK1) and peak in winter.');
      return answer({ title: 'What drives energy cost', obs, expl, next: ['Lowering PUE reduces cost most in high-price months.', 'Hedging and contracts are outside this data; involve finance for real cost.'], links: [['Open Energy & Cost', 'energy']], basis: 'Cost = facility MWh × synthetic market price index. Not Bulk\'s invoice: tariffs, taxes and hedges are excluded.' });
    } },
    { re: /pue|efficien|perform|how (is|are|did)|trend|status/, fn: s => {
      const a = agg(rows(s.months, s.sites)), ms = s.months.map(m => ({ m, ...agg(rows([m], s.sites)) })).filter(x => x.it);
      if (!ms.length) return answer({ title: 'No data', obs: [`No data for ${s.plabel}.`], expl: [], next: [] });
      const best = ms.reduce((x, y) => y.pue < x.pue ? y : x), worst = ms.reduce((x, y) => y.pue > x.pue ? y : x), above = ms.filter(x => x.pue > PUE_TARGET);
      return answer({ title: `Efficiency ${s.site ? `at ${s.site}` : 'across all sites'}`, obs: [`PUE <b>${a.pue.toFixed(3)}</b> in ${s.plabel} (target max. ${PUE_TARGET.toFixed(2)}). Cooling is ${(a.coolShare * 100).toFixed(1)}% of all energy.`, `Best month ${mLong(best.m)} (${best.pue.toFixed(3)}), worst ${mLong(worst.m)} (${worst.pue.toFixed(3)}). ${above.length} of ${ms.length} months are above target.`,
        ...(!s.site ? SITES.map(x => { const b = agg(rows(s.months, [x])); return b.it ? `${x}: PUE ${b.pue.toFixed(3)}` : null; }).filter(Boolean) : [])],
        expl: ['PUE usually rises in summer because cooling works harder; worse months that coincide with cooling flags point to a technical cause.'], next: ['Check months above target on the Sites page.'], links: [['Open Sites', 'sites', s.site || state.site]], basis: 'PUE = total ÷ IT energy, energy-weighted. Synthetic operating data.' });
    } },
    { re: /energy|consum|gwh|mwh|usage|use/, fn: s => {
      const a = agg(rows(s.months, s.sites));
      return answer({ title: `Energy use ${s.site ? `at ${s.site}` : ''}`, obs: [`${(a.total / 1000).toFixed(1)} GWh in ${s.plabel}: IT ${(a.it / a.total * 100).toFixed(0)}%, cooling ${(a.cool / a.total * 100).toFixed(0)}%, other ${(a.other / a.total * 100).toFixed(0)}%.`, ...(!s.site ? SITES.map(x => { const b = agg(rows(s.months, [x])); return b.total ? `${x}: ${(b.total / 1000).toFixed(1)} GWh (${(b.total / a.total * 100).toFixed(0)}%)` : null; }).filter(Boolean) : [])],
        expl: ['Rising energy mostly reflects more customer IT load; judge efficiency with PUE.'], next: [], links: [['Open Energy & Cost', 'energy']], basis: 'Monthly facility energy, synthetic.' });
    } },
    { re: /help|what can|how do i|hello|hi\b|hey/, fn: () => answer({ title: 'What I can help with', headings: { obs: 'Topics', next: 'Tip' }, obs: ['Cooling anomalies and their possible causes', 'Where the biggest savings are', 'Capacity and headroom', 'Comparison with Nordic peers', 'Energy cost and price drivers', 'PUE and energy use for a site or period'], expl: [], next: ['Name a site (N01, OS-IX, DK01, Oslo, Esbjerg…) or a period (2025, last month) in your question to change the focus. Otherwise I use the page\'s current context.'], basis: 'All answers are calculated from the loaded data with the same formulas as the dashboard.' }) },
  ];
  function respond(q) {
    const s = scope(q), hit = INTENTS.find(i => i.re.test(q.toLowerCase()));
    if (!hit) return answer({ title: 'I\'m not sure what you\'re asking', headings: { obs: 'What I can do', next: 'Try this' }, obs: ['I answer questions about energy, cooling, cost, capacity and peer benchmarks, using the dashboard\'s data.'], expl: [], next: ['Try one of the suggestions below, or rephrase with a keyword such as "cooling", "cost", "capacity" or "peers".'], suggest: true });
    const a = hit.fn(s); a.scope = `${s.site || 'All sites'} · ${s.plabel}`; return a;
  }
  function renderAnswer(a) {
    const H = a.headings || {};
    const list = (h, items, cls) => items && items.length ? `<div class="cp-sec ${cls}"><b>${h}</b><ul>${items.map(x => `<li>${x}</li>`).join('')}</ul></div>` : '';
    return `<div class="cp-msg bot"><span class="cp-av">${ic('bot', 16)}</span><div class="cp-body"><div class="cp-title">${a.title}</div>
      ${list(H.obs || 'Observations (from the data)', a.obs, 'obs')}${list('Possible explanations (not confirmed)', a.expl, 'expl')}${list(H.next || 'Suggested next steps', a.next, 'next')}
      ${a.links ? `<div class="cp-links">${a.links.map(([l, p, st]) => `<button class="btn ghost sm" data-cp-go="${p}" ${st ? `data-cp-site="${esc(st)}"` : ''}>${l} ${ic('arrow', 12)}</button>`).join('')}</div>` : ''}
      ${a.basis || a.scope ? `<div class="cp-basis">${a.scope ? `Scope: ${esc(a.scope)}. ` : ''}${a.basis || ''} Decision support, not a decision.</div>` : ''}
      ${a.suggest ? `<div class="cp-sugs">${SUGGESTIONS.map(x => `<button class="cp-sug" data-cp-q="${esc(x)}">${esc(x)}</button>`).join('')}</div>` : ''}</div></div>`;
  }

  // UI
  document.body.insertAdjacentHTML('beforeend', `
    <button class="cp-fab" id="cp-fab" aria-label="Ask Copilot" aria-keyshortcuts="${isMac ? 'Meta+J' : 'Control+J'}">${ic('bot', 20)}<span>Ask Copilot</span><kbd>${isMac ? '⌘ J' : 'Ctrl J'}</kbd></button>
    <aside class="cp-panel" id="cp-panel" role="dialog" aria-label="AI Copilot" hidden>
      <header class="cp-head"><span class="cp-hic">${ic('bot', 18)}</span><div><b>AI Copilot</b><div class="small">Decision support · workbook data · no external AI</div></div>
        <button class="cp-hbtn cp-exp" id="cp-expand" aria-label="Expand">${ic('expand', 17)}</button><button class="cp-hbtn" id="cp-close" aria-label="Close">${ic('x', 18)}</button></header>
      <div class="cp-log" id="cp-log">
        <div class="cp-msg bot"><span class="cp-av">${ic('bot', 16)}</span><div class="cp-body"><div class="cp-title">Explore what the energy, efficiency, anomaly, capacity and peer data shows.</div>
          <p class="cp-intro">Answers are calculated from the workbook with the same engine as the dashboards and separate observations from possible explanations. They support decisions rather than make them. No external AI service is used.</p></div></div>
        <div class="cp-sugs">${SUGGESTIONS.map(x => `<button class="cp-sug" data-cp-q="${esc(x)}">${esc(x)}</button>`).join('')}</div>
      </div>
      <form class="cp-foot" id="cp-form"><div class="cp-input"><input id="cp-q" autocomplete="off" maxlength="300" placeholder="Ask a question, e.g. Why did cooling energy increase?" aria-label="Ask a question"><button type="submit" class="cp-send" aria-label="Send">${ic('send', 17)}</button></div><div class="cp-ctx" id="cp-ctx"></div></form>
    </aside>`);
  const fab = document.getElementById('cp-fab'), panel = document.getElementById('cp-panel'), log = document.getElementById('cp-log'), qin = document.getElementById('cp-q');
  const openCp = () => { panel.hidden = false; document.body.classList.add('cp-open'); document.getElementById('cp-ctx').textContent = ctxLabel(); setTimeout(() => qin.focus(), 30); };
  const closeCp = () => { panel.hidden = true; document.body.classList.remove('cp-open'); fab.focus(); };
  function ask(q) {
    q = String(q || '').trim(); if (!q) return;
    log.querySelectorAll(':scope > .cp-sugs').forEach(el => el.remove());
    log.insertAdjacentHTML('beforeend', `<div class="cp-msg me"><div class="cp-body">${esc(q)}</div></div><div class="cp-msg bot typing"><span class="cp-av">${ic('bot', 16)}</span><div class="cp-body"><i></i><i></i><i></i></div></div>`);
    log.scrollTop = log.scrollHeight; qin.value = '';
    setTimeout(() => {
      log.querySelector('.typing')?.remove();
      let html; try { html = renderAnswer(respond(q)); } catch (err) { console.error(err); html = renderAnswer({ title: 'Something went wrong', obs: ['I could not calculate an answer from the current data.'], expl: [], next: [] }); }
      log.insertAdjacentHTML('beforeend', html);
      const msgs = log.querySelectorAll('.cp-msg.bot'); msgs[msgs.length - 1].scrollIntoView({ block: 'start', behavior: 'smooth' });
      if (log.querySelectorAll('.cp-msg.me').length === 1) log.insertAdjacentHTML('beforeend', `<div class="cp-sugs">${MORE.concat(SUGGESTIONS.filter(x => x !== q)).slice(0, 4).map(x => `<button class="cp-sug" data-cp-q="${esc(x)}">${esc(x)}</button>`).join('')}</div>`);
    }, 450);
  }
  fab.addEventListener('click', () => (panel.hidden ? openCp() : closeCp()));
  document.getElementById('cp-close').addEventListener('click', closeCp);
  document.getElementById('cp-expand').addEventListener('click', e => { const on = panel.classList.toggle('wide'); e.currentTarget.innerHTML = ic(on ? 'shrink' : 'expand', 17); e.currentTarget.setAttribute('aria-label', on ? 'Collapse' : 'Expand'); });
  document.getElementById('cp-form').addEventListener('submit', e => { e.preventDefault(); ask(qin.value); });
  panel.addEventListener('click', e => {
    const s = e.target.closest('[data-cp-q]'); if (s) { ask(s.dataset.cpQ); return; }
    const g = e.target.closest('[data-cp-go]');
    if (g) { if (g.dataset.cpSite) state.site = g.dataset.cpSite; if (matchMedia('(max-width: 640px)').matches) closeCp(); if (location.hash === '#' + g.dataset.cpGo) render(); else location.hash = g.dataset.cpGo; document.getElementById('cp-ctx').textContent = ctxLabel(); }
  });
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'j') { e.preventDefault(); panel.hidden ? openCp() : closeCp(); }
    else if (e.key === 'Escape' && !panel.hidden && !document.querySelector('.modal-bg')) closeCp();
  });

  /* ---------- router & events ---------- */
  // On phones the app shows only the core pages; the computer dashboard keeps all pages.
  const PHONE_PAGES = ['overview', 'sites', 'energy', 'insights'];
  const phoneMQ = matchMedia('(max-width: 640px)');
  phoneMQ.addEventListener('change', () => render());
  function render() {
    charts.forEach(c => c.destroy()); charts = [];
    const page = (location.hash || '#overview').slice(1);
    let key = pages[page] ? page : 'overview';
    if (phoneMQ.matches && !PHONE_PAGES.includes(key)) { key = 'overview'; history.replaceState(null, '', '#overview'); }
    document.querySelectorAll('[data-page]').forEach(a => a.classList.toggle('active', a.dataset.page === key));
    currentPage = key; hideInfo();
    pages[key]();
    decorateInfo(key);
    const cpc = document.getElementById('cp-ctx'); if (cpc) cpc.textContent = ctxLabel();
  }
  app.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.bind) { state[t.dataset.bind] = t.value; if (t.dataset.bind === 'period') store.set('period', t.value); render(); }
    if (t.dataset.check) { const d = store.get('actions', {}); d[t.dataset.check] = t.checked; store.set('actions', d); render(); }
  });
  app.addEventListener('click', e => {
    const t = e.target.closest('[data-tab],[data-go],[data-action],[data-download],[data-bench],[data-pick-site]');
    if (!t) return;
    if (t.dataset.tab) { const [k, v] = t.dataset.tab.split(':'); state[k] = v; render(); }
    if (t.dataset.go) { if (t.dataset.site) state.site = t.dataset.site; if (location.hash === '#' + t.dataset.go) render(); else location.hash = t.dataset.go; window.scrollTo(0, 0); }
    if (t.dataset.pickSite) { state.site = t.dataset.pickSite; render(); return; }
    if (t.dataset.bench) { toggleBench(t.dataset.bench); return; }
    if (t.dataset.download) downloadReport(allReports()[+t.dataset.download]);
    const a = t.dataset.action;
    if (a === 'new-report') newReportModal();
    if (a === 'details') detailsModal(t.dataset.month);
    if (a === 'bench-clear') { state.benchSel = []; store.set('benchSel', []); render(); }
    if (a === 'template') csv('Bulk_monthly_data_template.csv', MONTHLY_HEADER, monthlyCsvRows());
    if (a === 'remove-upload') { if (confirm('Remove this upload and recalculate the dashboard?')) { window.BulkUpload.remove(t.dataset.id); location.reload(); } }
    if (a === 'reset-uploads') { if (confirm('Remove all uploaded data and go back to the original Excel data?')) { window.BulkUpload.resetAll(); location.reload(); } }
    if (a === 'reset') { store.set('actions', {}); store.set('reports', []); render(); alert('Saved actions and reports cleared.'); }
  });
  document.querySelectorAll('[data-icon]').forEach(el => { el.innerHTML = ic(el.dataset.icon); });
  window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
  if (!PERIODS[state.period]) state.period = 'ytd';
  render();
  const flash = sessionStorage.getItem('bulk2.flash');
  if (flash) {
    sessionStorage.removeItem('bulk2.flash');
    const t = document.getElementById('pwa-toast');
    if (t) { t.innerHTML = `${ic('check', 15)} ${esc(flash)}`; t.hidden = false; setTimeout(() => { t.hidden = true; }, 4000); }
  }
})();
