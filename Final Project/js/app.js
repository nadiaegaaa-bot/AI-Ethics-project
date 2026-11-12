/* Bulk Energy Intelligence dashboard. Every number is computed from window.BULK_DATA (js/data.js), generated from the Excel workbook. */
(() => {
  const ICONS = {"home": "<path d=\"M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z\"/>", "sites": "<rect x=\"3\" y=\"7\" width=\"7\" height=\"14\"/><rect x=\"10\" y=\"3\" width=\"11\" height=\"18\"/><path d=\"M13 7h2M17 7h1M13 11h2M17 11h1M13 15h2M17 15h1\"/>", "bolt": "<path d=\"M13 2L4 14h7l-1 8 9-12h-7z\"/>", "spark": "<path d=\"M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8zM19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z\"/>", "db": "<ellipse cx=\"12\" cy=\"5\" rx=\"8\" ry=\"3\"/><path d=\"M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3\"/>", "chart": "<path d=\"M3 3v18h18M8 17V11M13 17V7M18 17v-4\"/>", "file": "<path d=\"M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z\"/><path d=\"M14 3v5h5M9 13h6M9 17h6\"/>", "gear": "<circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z\"/>", "warn": "<path d=\"M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0zM12 9v4M12 17h.01\"/>", "bulb": "<path d=\"M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.1V17h6v-.2c0-.8.4-1.6 1-2.1A7 7 0 0 0 12 2z\"/>", "check": "<path d=\"M20 6L9 17l-5-5\"/>", "info": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M12 16v-4M12 8h.01\"/>", "target": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><circle cx=\"12\" cy=\"12\" r=\"5\"/><circle cx=\"12\" cy=\"12\" r=\"1\"/>", "leaf": "<path d=\"M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.5 19 2c1 2 2 4.2 2 8 0 5.5-4.8 10-10 10zM2 21c0-3 1.9-5.4 5.1-6\"/>", "link": "<path d=\"M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7\"/>", "down": "<path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3\"/>", "up": "<path d=\"M3 17l6-6 4 4 8-8M15 7h6v6\"/>", "dn": "<path d=\"M3 7l6 6 4-4 8 8M15 17h6v-6\"/>", "chev": "<path d=\"M9 18l6-6-6-6\"/>", "arrow": "<path d=\"M5 12h14M13 6l6 6-6 6\"/>", "db2": "<ellipse cx=\"12\" cy=\"5\" rx=\"8\" ry=\"3\"/><path d=\"M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 11v8c0 1.7 3.6 3 8 3s8-1.3 8-3v-8\"/>", "book": "<path d=\"M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z\"/><path d=\"M4 19.5A2.5 2.5 0 0 0 6.5 22H20v-5\"/>", "ext": "<path d=\"M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14L21 3\"/>"};
  const ic = (n, s = 15) => `<svg class="icon" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICONS[n] || ''}</svg>`;

  const D = window.BULK_DATA;
  const C = { purple: '#7a5cf5', blue: '#4fa3f0', other: '#d9dbe3', green: '#3ecf72', amber: '#f6b93b', red: '#ff5c5c', muted: '#9aa0b8', grid: '#23273b' };
  const SITE_INFO = {
    'N01': { location: 'Vennesla (Kristiansand)', country: 'Norway' },
    'OS-IX': { location: 'Oslo', country: 'Norway' },
    'DK01': { location: 'Esbjerg', country: 'Denmark' },
  };
  const SITES = ['OS-IX', 'N01', 'DK01'].filter(s => D.monthly.some(r => r.site === s));
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
  const badge = (t = 'Synthetic data') => `<span class="badge" title="Values are simulated in the source workbook, not measured Bulk data">${t}</span>`;
  const periodSelect = () => `<select data-bind="period">${Object.entries(PERIODS).map(([k, l]) => `<option value="${k}" ${k === state.period ? 'selected' : ''}>${l}</option>`).join('')}</select>`;
  const siteSelect = () => `<select data-bind="site">${SITES.map(s => `<option ${s === state.site ? 'selected' : ''}>${s}</option>`).join('')}</select>`;
  const rangeLabel = (ms = periodMonths().cur) => ms.length === 1 ? mLong(ms[0]) : `${mLong(ms[0])} – ${mLong(ms.at(-1))}`;
  const head = (title, sub, controls = '') => `<div class="head"><div><h1>${title}</h1><div class="sub">${sub}</div></div><div class="controls">${controls}</div></div>`;
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

  pages.overview = () => {
    const a = cur(), p = prev(), ms = periodMonths().cur;
    const flags = flagged(ms), latestAny = flagged()[0];
    const excess = flags.reduce((t, r) => t + Math.max(r.residualMWh, 0), 0);
    const insights = flags.length
      ? flags.slice(0, 3).map(r => insightCard(r)).join('') +
        `<div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('bulb')}</span><div><b>Potential saving</b>Investigate cooling system performance at ${[...new Set(flags.map(f => f.site))].join(', ')}. Excess cooling vs. model: ${excess.toFixed(0)} MWh (${(excess / a.total * 100).toFixed(2)}% of total energy), about ${nokM(flags.reduce((t, r) => t + excessNOK(r), 0))} at the spot index.</div></div>`
      : `<div class="insight"><span style="color:${C.green}">${ic('check')}</span><div><b>No cooling review flags in this period</b>All site-months are within ±${THRESH * 100}% of expected cooling.</div></div>` +
        (latestAny ? insightCard(latestAny) : '');
    app.innerHTML = head('Overview', `Energy performance across all sites · ${rangeLabel()} ${badge()}`, periodSelect()) + `
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
        ${note(`Dashed line: Bulk's new-build design PUE target (≤ ${PUE_TARGET.toFixed(2)}), from the Annual Report 2025. It is a target, not a measured value.`)}`;
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
        <div class="insight"><span style="color:${C.blue}">${ic('info')}</span><div><b>PUE vs. design target</b>PUE ${a.pue.toFixed(3)} vs. target ≤ ${PUE_TARGET.toFixed(2)}: ${a.pue <= PUE_TARGET ? 'within target' : `${((a.pue - PUE_TARGET) / PUE_TARGET * 100).toFixed(1)}% above target`}.</div></div></div></div>`;
    }
    app.innerHTML = head(`Site: ${s}`, `${esc(info.location)} · ${esc(zone(s))} · ${rangeLabel()} ${badge()}`, periodSelect() + siteSelect()) +
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
    app.innerHTML = head('AI Insights', `AI-powered analysis of energy use, efficiency and capacity ${badge()}`, siteSelect() + periodSelect()) + banner + impact +
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
      <div class="grid g3">${SITES.map(s => card(lastRows.find(r => r.site === s))).join('')}</div>
      <div class="card" style="margin-bottom:12px"><h3>Average IT load by site (MW)</h3><div class="chart"><canvas id="c1"></canvas></div></div>
      ${note(`N01's ${cap} MW is the operating IT capacity reported for 2025 (Menon Economics, Aug 2026). It is capacity, not observed load. The workbook has no verified capacity for OS-IX or DK01, so no utilisation is shown for them.`)}`;
    const colors = { 'OS-IX': C.blue, 'N01': C.purple, 'DK01': C.green };
    mk('c1', { type: 'line', data: { labels: MONTHS.map(mShort), datasets: SITES.map(s => ({ label: s, data: rows(MONTHS, [s]).map(r => r.itLoadMW), borderColor: colors[s], backgroundColor: colors[s], pointRadius: 0, borderWidth: 2 })) },
      options: opts({ scales: { x: noGridX, y: { beginAtZero: true } }, plugins: { legend: { position: 'bottom' }, overlays: { hlines: cap ? [{ y: cap, label: `N01 capacity ${cap} MW` }] : [] } } }) });
  };

  pages.benchmarking = () => {
    const b2025 = agg(rows(MONTHS.filter(m => m.startsWith('2025'))));
    const numeric = D.benchmarks.filter(b => typeof b.pue === 'number' && b.status === 'VERIFIED');
    // Independent peers: skip EcoDataCenter Falun (same operator as portfolio row).
    const peers = numeric.filter(b => !/Falun/.test(b.facility));
    const mkt = peers.reduce((t, b) => t + b.pue, 0) / peers.length;
    const peerNames = [...new Set(D.peerMonthly.map(r => r.peer))];
    const peerMonths = [...new Set(D.peerMonthly.map(r => r.month))].sort();
    let left, right;
    if (state.benchTab === 'reported') {
      left = ['Reported annual PUE (public reports)', 'c1'];
      right = `<div class="rank me"><span>Bulk 2025 (synthetic)</span><b>${b2025.pue.toFixed(3)}</b></div>` +
        numeric.map(b => `<div class="rank"><span>${esc(b.operator)} · ${esc(b.facility)} <span class="small">(${b.year})</span></span><b>${b.pue.toFixed(2)}</b></div>`).join('') +
        `<div class="rank"><span>Peer average (excl. Falun)</span><b>${mkt.toFixed(2)}</b></div>`;
    } else if (state.benchTab === 'monthly') {
      left = ['Monthly PUE scenarios 2025', 'c1'];
      right = peerNames.map(n => { const r = D.peerMonthly.filter(x => x.peer === n); const v = r.reduce((t, x) => t + x.totalMWh, 0) / r.reduce((t, x) => t + x.itMWh, 0); return `<div class="rank"><span>${esc(n)}</span><b>${v.toFixed(3)}</b></div>`; }).join('');
      right = `<div class="rank me"><span>Bulk (all sites)</span><b>${b2025.pue.toFixed(3)}</b></div>` + right;
    } else {
      left = ['Market price index by zone (NOK/MWh)', 'c1'];
      right = SITES.map(s => { const r = rows(MONTHS, [s]); const avg = r.reduce((t, x) => t + x.priceNOKMWh, 0) / r.length; return `<div class="rank"><span>${esc(zone(s))} (${s})</span><b>${avg.toFixed(0)}</b></div>`; }).join('');
    }
    const n01 = rows(MONTHS, ['N01']);
    const hi = n01.reduce((m, r) => r.priceNOKMWh > m.priceNOKMWh ? r : m, n01[0]);
    const avgP = n01.reduce((t, r) => t + r.priceNOKMWh, 0) / n01.length;
    app.innerHTML = head('Benchmarking', `Compare Bulk's sites with selected competitors and market context`, '') +
      tabs('benchTab', [['reported', 'Reported PUE (verified)'], ['monthly', 'Monthly scenarios (synthetic)'], ['price', 'Electricity price']]) + `
      <div class="grid g21"><div class="card"><h3>${left[0]}</h3><div class="chart tall"><canvas id="c1"></canvas></div></div><div class="card"><h3>${state.benchTab === 'price' ? 'Average price (all months)' : 'Average PUE'}</h3>${right}</div></div>
      <div class="grid g21"><div class="card"><h3>Electricity price N01 (NOK/MWh)</h3><div class="chart"><canvas id="c2"></canvas></div></div>
      <div class="card"><h3><span style="color:${C.green}">${ic('info')}</span> Key takeaway</h3><p class="takeaway">
        Bulk's 2025 PUE in the synthetic data (${b2025.pue.toFixed(3)}) is ${((mkt - b2025.pue) / mkt * 100).toFixed(1)}% lower than the average reported by Nordic peers (${mkt.toFixed(2)}). Because Bulk's figure is simulated and the peers use different years, scopes and methods, this is an illustration, not a ranking.
        N01's price index peaked in ${mLong(hi.month)} (${hi.priceNOKMWh.toFixed(0)} NOK/MWh), ${((hi.priceNOKMWh - avgP) / avgP * 100).toFixed(0)}% above its average.</p></div></div>
      ${state.benchTab === 'reported' ? `<div class="card" style="margin-bottom:12px"><h3>Comparability notes</h3><table><thead><tr><th>Operator</th><th>Scope</th><th>Status</th><th>Caution</th><th>Source</th></tr></thead><tbody>
        ${D.benchmarks.map(b => `<tr><td><b>${esc(b.operator)}</b><div class="small">${esc(b.facility)}</div></td><td>${esc(b.pueType)}</td><td>${esc(b.status)}</td><td>${esc(b.cautions)}</td><td>${/^https?:/.test(b.source || '') ? `<a href="${esc(b.source)}" target="_blank" rel="noopener">Link ${ic('ext', 11)}</a>` : esc(b.source)}</td></tr>`).join('')}
      </tbody></table></div>` : ''}
      ${state.benchTab === 'monthly' ? note('Monthly peer values are synthetic scenarios anchored only to each peer\'s published annual PUE. They are not observed peer operations.') : ''}`;
    if (state.benchTab === 'reported') {
      const labels = ['Bulk 2025*', ...numeric.map(b => `${b.operator}${/portfolio/i.test(b.facility) ? '' : ' ' + b.facility.split(' ')[0]}`)];
      mk('c1', { type: 'bar', data: { labels, datasets: [{ data: [+b2025.pue.toFixed(3), ...numeric.map(b => b.pue)], backgroundColor: [C.green, ...numeric.map(() => '#4a4f6a')], maxBarThickness: 60 }] },
        options: opts({ plugins: { legend: { display: false }, overlays: { hlines: [{ y: PUE_TARGET, label: `Bulk design target ${PUE_TARGET.toFixed(2)}` }] }, tooltip: { callbacks: { afterLabel: c => c.dataIndex === 0 ? 'Synthetic, not reported' : `${numeric[c.dataIndex - 1].year} · ${numeric[c.dataIndex - 1].pueType}` } } }, scales: { x: noGridX, y: { min: 1.0, max: 1.45 } } }) });
    } else if (state.benchTab === 'monthly') {
      const bulkM = peerMonths.map(m => +agg(rows([m])).pue.toFixed(4));
      const shades = ['#e6e7ee', '#b8bbd0', '#8f93ad', '#6d7190', '#55597a'];
      mk('c1', { type: 'line', data: { labels: peerMonths.map(mShort), datasets: [
        { label: 'Bulk (avg.)', data: bulkM, borderColor: C.green, backgroundColor: C.green, borderWidth: 2.5, pointRadius: 3 },
        ...peerNames.map((n, i) => ({ label: n, data: peerMonths.map(m => (D.peerMonthly.find(r => r.peer === n && r.month === m) || {}).pue), borderColor: shades[i], borderDash: i % 2 ? [5, 4] : [], pointRadius: 0, borderWidth: 1 }))] },
        options: opts({ scales: { x: noGridX, y: { suggestedMin: 1.15, suggestedMax: 1.42 } } }) });
    } else {
      const colors = { 'OS-IX': C.blue, 'N01': '#2f9e5f', 'DK01': C.amber };
      mk('c1', { type: 'line', data: { labels: MONTHS.map(mShort), datasets: SITES.map(s => ({ label: zone(s), data: rows(MONTHS, [s]).map(r => r.priceNOKMWh), borderColor: colors[s], backgroundColor: colors[s], pointRadius: 0, borderWidth: 2 })) },
        options: opts({ scales: { x: noGridX, y: { beginAtZero: true } } }) });
    }
    mk('c2', { type: 'line', data: { labels: n01.map(r => mShort(r.month)), datasets: [{ label: 'NOK/MWh', data: n01.map(r => r.priceNOKMWh), borderColor: '#2f9e5f', backgroundColor: 'rgba(47,158,95,0.35)', fill: true, pointRadius: 0, borderWidth: 1.5 }] },
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
    app.innerHTML = head('Reports', 'Download generated reports', '<button class="btn" data-action="new-report">New report</button>') +
      `<div class="card">${allReports().map((r, i) => `<div class="report"><span class="ic">${ic('file', 18)}</span><div><b>${esc(r.title)}</b><div class="small">${fmtDate(r.date)} · ${r.type === 'XLSX' ? 'XLSX (CSV)' : r.type}</div></div><button class="btn ghost" data-download="${i}">${ic('down', 13)} Download</button></div>`).join('')}</div>
      ${note('PDF reports open in a print view. Choose "Save as PDF" in the print dialog. Data exports download as CSV, which opens in Excel.')}`;
  };

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
      csv(r.title.replace(/[^\w]+/g, '_') + '.csv', ['Month', 'Site', 'Zone', 'IT avg load (MW)', 'IT energy (MWh)', 'Cooling (MWh)', 'Other (MWh)', 'Total (MWh)', 'PUE', 'Outdoor °C', 'Price NOK/MWh', 'Spot-index cost NOK', 'Expected cooling (MWh)', 'Residual %', 'Alert', 'Provenance'],
        D.monthly.map(m => [m.month, m.site, m.zone, m.itLoadMW, m.itMWh, m.coolingMWh, m.otherMWh, m.totalMWh, m.pue, m.tempC, m.priceNOKMWh, m.costNOK, m.expectedCoolingMWh, m.residualPct, m.alert, m.provenance]));
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
      <p class="warn">Prototype report. All operational values are SYNTHETIC (source: ${esc(D.source)}). Cost is facility MWh × a synthetic price index, not Bulk's invoice.</p>
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
      `<div class="card">${body}</div>`;
  };

  pages.settings = () => {
    app.innerHTML = head('Settings', 'Data and preferences') + `
      <div class="card" style="font-size:13px;line-height:1.7"><h3>Data file</h3>
        <p>Loaded from <code>js/data.js</code>, generated from <b>${esc(D.source)}</b> on ${esc(D.generated)}: ${D.monthly.length} site-months (${SITES.length} sites, ${mLong(MONTHS[0])} – ${mLong(LAST)}), ${D.benchmarks.length} competitor benchmarks, ${D.peerMonthly.length} peer scenario months, ${D.facts.length} company facts.</p>
        <p>To update after editing the Excel file, run <code>python3 tools/excel_to_data.py</code> in the <code>Final Project</code> folder.</p>
        <p>Saved in this browser only: the period you selected, ticked AI actions and reports you created.</p>
        <button class="btn ghost" data-action="reset">Reset saved actions and reports</button></div>`;
  };

  /* ---------- router & events ---------- */
  function render() {
    charts.forEach(c => c.destroy()); charts = [];
    const page = (location.hash || '#overview').slice(1);
    const key = pages[page] ? page : 'overview';
    document.querySelectorAll('[data-page]').forEach(a => a.classList.toggle('active', a.dataset.page === key));
    pages[key]();
  }
  app.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.bind) { state[t.dataset.bind] = t.value; if (t.dataset.bind === 'period') store.set('period', t.value); render(); }
    if (t.dataset.check) { const d = store.get('actions', {}); d[t.dataset.check] = t.checked; store.set('actions', d); render(); }
  });
  app.addEventListener('click', e => {
    const t = e.target.closest('[data-tab],[data-go],[data-action],[data-download]');
    if (!t) return;
    if (t.dataset.tab) { const [k, v] = t.dataset.tab.split(':'); state[k] = v; render(); }
    if (t.dataset.go) { if (t.dataset.site) state.site = t.dataset.site; if (location.hash === '#' + t.dataset.go) render(); else location.hash = t.dataset.go; window.scrollTo(0, 0); }
    if (t.dataset.download) downloadReport(allReports()[+t.dataset.download]);
    const a = t.dataset.action;
    if (a === 'new-report') newReportModal();
    if (a === 'details') detailsModal(t.dataset.month);
    if (a === 'reset') { store.set('actions', {}); store.set('reports', []); render(); alert('Saved actions and reports cleared.'); }
  });
  document.querySelectorAll('[data-icon]').forEach(el => { el.innerHTML = ic(el.dataset.icon); });
  window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
  if (!PERIODS[state.period]) state.period = 'ytd';
  render();
})();
