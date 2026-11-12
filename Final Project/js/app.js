/* Bulk Energy Intelligence dashboard. All numbers are computed from window.BULK_DATA (js/data.js). */
(() => {
  const D = window.BULK_DATA;
  const C = { purple: '#7a5cf5', blue: '#4fa3f0', other: '#d9dbe3', green: '#3ecf72', amber: '#f6b93b', red: '#ff5c5c', muted: '#9aa0b8', grid: '#23273b' };
  const MONTHS = [...new Set(D.monthly.map(r => r.month))].sort();
  const SITES = D.sites.map(s => s.id);
  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const mLabel = m => MON[+m.slice(5) - 1];
  const mLong = m => `${MON[+m.slice(5) - 1]} ${m.slice(0, 4)}`;
  const app = document.getElementById('app');
  const ICONS = {"home": "<path d=\"M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z\"/>", "sites": "<rect x=\"3\" y=\"7\" width=\"7\" height=\"14\"/><rect x=\"10\" y=\"3\" width=\"11\" height=\"18\"/><path d=\"M13 7h2M17 7h1M13 11h2M17 11h1M13 15h2M17 15h1\"/>", "bolt": "<path d=\"M13 2L4 14h7l-1 8 9-12h-7z\"/>", "spark": "<path d=\"M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8zM19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z\"/>", "db": "<ellipse cx=\"12\" cy=\"5\" rx=\"8\" ry=\"3\"/><path d=\"M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3\"/>", "chart": "<path d=\"M3 3v18h18M8 17V11M13 17V7M18 17v-4\"/>", "file": "<path d=\"M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z\"/><path d=\"M14 3v5h5M9 13h6M9 17h6\"/>", "gear": "<circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z\"/>", "warn": "<path d=\"M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0zM12 9v4M12 17h.01\"/>", "bulb": "<path d=\"M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.1V17h6v-.2c0-.8.4-1.6 1-2.1A7 7 0 0 0 12 2z\"/>", "check": "<path d=\"M20 6L9 17l-5-5\"/>", "info": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M12 16v-4M12 8h.01\"/>", "target": "<circle cx=\"12\" cy=\"12\" r=\"9\"/><circle cx=\"12\" cy=\"12\" r=\"5\"/><circle cx=\"12\" cy=\"12\" r=\"1\"/>", "leaf": "<path d=\"M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.5 19 2c1 2 2 4.2 2 8 0 5.5-4.8 10-10 10zM2 21c0-3 1.9-5.4 5.1-6\"/>", "link": "<path d=\"M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7\"/>", "down": "<path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3\"/>", "up": "<path d=\"M3 17l6-6 4 4 8-8M15 7h6v6\"/>", "dn": "<path d=\"M3 7l6 6 4-4 8 8M15 17h6v-6\"/>", "chev": "<path d=\"M9 18l6-6-6-6\"/>", "arrow": "<path d=\"M5 12h14M13 6l6 6-6 6\"/>"};
  const ic = (n, s = 15) => `<svg class="icon" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICONS[n]}</svg>`;
  const store = {
    get: (k, d) => { try { return JSON.parse(localStorage.getItem('bulk.' + k)) ?? d; } catch { return d; } },
    set: (k, v) => localStorage.setItem('bulk.' + k, JSON.stringify(v)),
  };
  const state = { period: store.get('period', '6m'), site: 'N01', siteTab: 'performance', benchTab: 'efficiency' };
  let charts = [];

  Chart.defaults.color = C.muted;
  Chart.defaults.font.family = 'Manrope, system-ui, sans-serif';
  Chart.defaults.font.size = 11;
  Chart.defaults.borderColor = C.grid;
  Chart.defaults.plugins.legend.labels.boxWidth = 10;

  // Small overlay plugin: dashed horizontal target line and shaded x-range box.
  Chart.register({
    id: 'overlays',
    beforeDatasetsDraw(chart, _a, o) {
      const { ctx, chartArea: ca, scales: { x, y } } = chart;
      if (o.box) {
        const x0 = x.getPixelForValue(o.box.from) - (x.getPixelForValue(1) - x.getPixelForValue(0)) / 2;
        const x1 = x.getPixelForValue(o.box.to) + (x.getPixelForValue(1) - x.getPixelForValue(0)) / 2;
        ctx.save(); ctx.fillStyle = 'rgba(122,92,245,0.18)'; ctx.fillRect(x0, ca.top, x1 - x0, ca.bottom - ca.top);
        ctx.fillStyle = '#fff'; ctx.font = '11px Manrope, sans-serif'; ctx.textAlign = 'center'; ctx.fillText(o.box.label, (x0 + x1) / 2, ca.top + 14); ctx.restore();
      }
      if (o.hline) {
        const py = y.getPixelForValue(o.hline.y);
        ctx.save(); ctx.strokeStyle = '#666b85'; ctx.setLineDash([5, 5]); ctx.beginPath(); ctx.moveTo(ca.left, py); ctx.lineTo(ca.right, py); ctx.stroke();
        ctx.setLineDash([]); ctx.fillStyle = C.muted; ctx.font = '10px Manrope, sans-serif'; ctx.textAlign = 'center';
        ctx.fillStyle = '#141726'; const w = ctx.measureText(o.hline.label).width + 10; ctx.fillRect((ca.left + ca.right) / 2 - w / 2, py - 7, w, 14);
        ctx.fillStyle = C.muted; ctx.fillText(o.hline.label, (ca.left + ca.right) / 2, py + 3); ctx.restore();
      }
    },
  });

  /* ---------- data helpers ---------- */
  const PERIODS = { '1m': ['Last 30 days', 1], '6m': ['Last 6 months', 6], '12m': ['Last 12 months', 12] };
  const periodMonths = (p = state.period, offset = 0) => {
    const n = PERIODS[p][1], end = MONTHS.length - offset * n;
    return end - n < 0 ? [] : MONTHS.slice(end - n, end);
  };
  const rows = (months, sites = SITES) => D.monthly.filter(r => months.includes(r.month) && sites.includes(r.site));
  const agg = rs => {
    const a = rs.reduce((t, r) => ({ it: t.it + r.itMWh, cool: t.cool + r.coolingMWh, other: t.other + r.otherMWh, cost: t.cost + r.costNOK }), { it: 0, cool: 0, other: 0, cost: 0 });
    a.total = a.it + a.cool + a.other; a.pue = a.it ? a.total / a.it : 0; a.price = a.total ? a.cost / (a.total * 1000) : 0;
    return a;
  };
  const cur = (sites) => agg(rows(periodMonths(), sites));
  const prev = (sites) => { const m = periodMonths(state.period, 1); return m.length ? agg(rows(m, sites)) : null; };
  const site = id => D.sites.find(s => s.id === id);

  /* ---------- format helpers ---------- */
  const gwh = mwh => `${(mwh / 1000).toFixed(1)} GWh`;
  const nokM = n => `NOK ${(n / 1e6).toFixed(1)}M`;
  const pct = (a, b) => (b ? (a - b) / b * 100 : null);
  const delta = (now, before, lowerIsGood = false) => {
    const d = before == null ? null : pct(now, before);
    if (d == null) return '<div class="delta">No previous period</div>';
    const up = d >= 0, good = lowerIsGood ? !up : up;
    return `<div class="delta"><span class="${good ? 'up-good' : 'up-bad'}">${ic(up ? 'up' : 'dn', 12)} ${up ? '+' : ''}${d.toFixed(1)}%</span> vs. previous</div>`;
  };
  const kpi = (label, value, foot) => `<div class="card kpi"><div class="label">${label}</div><div class="value">${value}</div>${foot}</div>`;
  const periodSelect = () => `<select data-bind="period">${Object.entries(PERIODS).map(([k, [l]]) => `<option value="${k}" ${k === state.period ? 'selected' : ''}>${l}</option>`).join('')}</select>`;
  const siteSelect = () => `<select data-bind="site">${SITES.map(s => `<option ${s === state.site ? 'selected' : ''}>${s}</option>`).join('')}</select>`;
  const rangeLabel = () => { const m = periodMonths(); return m.length === 1 ? mLong(m[0]) : `${mLong(m[0])} – ${mLong(m[m.length - 1])}`; };
  const head = (title, sub, controls = '') => `<div class="head"><div><h1>${title}</h1><div class="sub">${sub}</div></div><div class="controls">${controls}</div></div>`;
  const tabs = (key, items) => `<div class="tabs">${items.map(([k, l]) => `<button data-tab="${key}:${k}" class="${state[key] === k ? 'active' : ''}">${l}</button>`).join('')}</div>`;

  /* ---------- chart helpers ---------- */
  const mk = (id, cfg) => { const el = document.getElementById(id); if (el) charts.push(new Chart(el, cfg)); };
  const opts = (extra = {}) => ({ responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom' } }, ...extra });

  /* ---------- AI analysis (rule-based on the data) ---------- */
  function anomalies(siteId) {
    const out = [];
    D.insights.filter(i => !siteId || i.site === siteId).forEach(i => out.push(i));
    // Cooling share check per site: flag if cooling share in latest month is >10% above its 12-month average.
    SITES.filter(s => !siteId || s === siteId).forEach(s => {
      const last = agg(rows([MONTHS.at(-1)], [s])), year = agg(rows(MONTHS, [s]));
      const shareNow = last.cool / last.total, shareAvg = year.cool / year.total;
      if (shareNow > shareAvg * 1.1 && !out.some(o => o.site === s))
        out.push({ site: s, severity: 'warning', title: `Higher cooling share at ${s}`, summary: `Cooling was ${(shareNow * 100).toFixed(0)}% of energy in ${mLong(MONTHS.at(-1))} vs. ${(shareAvg * 100).toFixed(0)}% on average.` });
    });
    return out;
  }
  const coolingExcess = () => {
    const win = D.coolingDailyN01.filter(d => d.actualMWh > d.expectedMWh * 1.05);
    const act = win.reduce((t, d) => t + d.actualMWh, 0), exp = win.reduce((t, d) => t + d.expectedMWh, 0);
    return { win, pct: exp ? (act - exp) / exp * 100 : 0 };
  };

  /* ---------- pages ---------- */
  const pages = {};

  pages.overview = () => {
    const a = cur(), p = prev();
    const ins = D.insights[0], ex = coolingExcess();
    app.innerHTML = head('Overview', 'Energy performance across all sites', periodSelect()) + `
      <div class="grid g4">
        ${kpi('Average PUE', a.pue.toFixed(2), delta(a.pue, p?.pue, true))}
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total))}
        ${kpi('Cooling Energy', gwh(a.cool), `<div class="delta">${(a.cool / a.total * 100).toFixed(0)}% of total</div>`)}
        ${kpi('Estimated Energy Cost', nokM(a.cost), delta(a.cost, p?.cost, true))}
      </div>
      <div class="grid g21">
        <div class="card"><h3>Energy consumption by site</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3><span style="color:${C.purple}">${ic('spark')}</span> AI insights</h3>
          <div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('warn')}</span><div><b>Higher cooling energy at ${ins.site}</b>+${ex.pct.toFixed(0)}% vs. expected. Not explained by weather or IT load.</div><span class="chev">${ic('chev')}</span></div>
          <div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('bulb')}</span><div><b>Potential saving</b>Investigate cooling system performance at ${ins.site}. Estimated impact: up to −${(ins.savingGWh * 1000 / a.total * 100).toFixed(0)}% total energy.</div></div>
          ${anomalies().filter(i => i !== ins).map(i => `<div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('warn')}</span><div><b>${i.title}</b>${i.summary}</div></div>`).join('')}
        </div>
      </div>`;
    const by = SITES.map(s => cur([s]));
    mk('c1', { type: 'bar', data: { labels: SITES, datasets: [
      { label: 'IT energy', data: by.map(b => b.it / 1000), backgroundColor: C.purple, stack: 's' },
      { label: 'Cooling energy', data: by.map(b => b.cool / 1000), backgroundColor: C.blue, stack: 's' },
      { label: 'Other', data: by.map(b => b.other / 1000), backgroundColor: C.other, stack: 's' }] },
      options: opts({ scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, ticks: { callback: v => `${v} GWh` } } }, datasets: { bar: { maxBarThickness: 120 } } }) });
  };

  pages.sites = () => {
    const s = state.site, a = cur([s]), p = prev([s]), ms = periodMonths();
    let body = '';
    if (state.siteTab === 'performance') {
      body = `<div class="grid g4">
        ${kpi('PUE', a.pue.toFixed(2), delta(a.pue, p?.pue, true))}
        ${kpi('IT Energy', gwh(a.it), delta(a.it, p?.it))}
        ${kpi('Cooling Energy', gwh(a.cool), delta(a.cool, p?.cool, true))}
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total, true))}</div>
        <div class="grid g2"><div class="card"><h3>PUE trend</h3><div class="chart"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3>Energy breakdown (${mLabel(ms.at(-1))})</h3><div class="donut-wrap"><div class="chart"><canvas id="c2"></canvas><div class="center-label" id="dl"></div></div><div class="legend" id="lg"></div></div></div></div>`;
    } else if (state.siteTab === 'breakdown') {
      body = `<div class="card"><h3>Monthly energy breakdown (GWh)</h3><div class="chart tall"><canvas id="c3"></canvas></div></div>`;
    } else if (state.siteTab === 'cost') {
      body = `<div class="grid g4">${kpi('Energy cost', nokM(a.cost), delta(a.cost, p?.cost, true))}${kpi('Avg. price', `${a.price.toFixed(2)} NOK/kWh`, delta(a.price, p?.price, true))}${kpi('Location', site(s).location, `<div class="delta">${site(s).country}</div>`)}${kpi('Renewable share', `${site(s).renewablePct}%`, '<div class="delta">Hydro &amp; wind PPAs</div>')}</div>
        <div class="card"><h3>Monthly cost (NOK M) and price (NOK/kWh)</h3><div class="chart tall"><canvas id="c4"></canvas></div></div>`;
    } else {
      const list = anomalies(s);
      body = `<div class="card"><h3><span style="color:${C.purple}">${ic('spark')}</span> AI analysis for ${s}</h3>${list.length ? list.map(i => `<div class="insight" data-go="insights"><span style="color:${C.amber}">${ic('warn')}</span><div><b>${i.title}</b>${i.summary}</div><span class="chev">${ic('chev')}</span></div>`).join('') : `<div class="insight"><span style="color:${C.green}">${ic('check')}</span><div><b>No anomalies detected</b>Energy use at ${s} is in line with the expected model for ${rangeLabel()}.</div></div>`}
        <div class="insight"><span style="color:${C.blue}">${ic('info')}</span><div><b>PUE vs. target</b>PUE ${a.pue.toFixed(2)} vs. target ${D.pueTarget.toFixed(2)}: ${a.pue <= D.pueTarget ? 'on target' : `${((a.pue - D.pueTarget) / D.pueTarget * 100).toFixed(1)}% above target`}.</div></div></div>`;
    }
    app.innerHTML = head(`Site: ${s}`, `${site(s).location} · ${rangeLabel()}`, periodSelect() + siteSelect()) +
      tabs('siteTab', [['performance', 'Performance'], ['breakdown', 'Energy breakdown'], ['cost', 'Cost &amp; prices'], ['analysis', 'AI analysis']]) + body;
    const mrows = ms.map(m => agg(rows([m], [s])));
    mk('c1', { type: 'line', data: { labels: ms.map(mLabel), datasets: [{ label: 'PUE', data: mrows.map(r => +r.pue.toFixed(3)), borderColor: C.purple, backgroundColor: C.purple, tension: 0, pointRadius: 3 }] },
      options: opts({ plugins: { legend: { display: false }, overlays: { hline: { y: D.pueTarget, label: `Target ≤ ${D.pueTarget.toFixed(2)}` } } },
        scales: { x: { grid: { display: false } }, y: { suggestedMin: D.pueTarget - 0.05, suggestedMax: Math.max(...mrows.map(r => r.pue)) + 0.04 } } }) });
    const last = mrows.at(-1);
    if (document.getElementById('c2')) {
      const parts = [['IT energy', last.it, C.purple], ['Cooling energy', last.cool, C.blue], ['Other', last.other, C.other]];
      document.getElementById('dl').innerHTML = `<b>${(last.total / 1000).toFixed(1)}</b><span class="small">GWh</span>`;
      document.getElementById('lg').innerHTML = parts.map(([l, v, c]) => `<div style="display:flex"><span><span class="dot" style="background:${c}"></span>${l}</span><b>${(v / last.total * 100).toFixed(0)}%</b></div>`).join('');
      mk('c2', { type: 'doughnut', data: { labels: parts.map(p => p[0]), datasets: [{ data: parts.map(p => +(p[1] / 1000).toFixed(2)), backgroundColor: parts.map(p => p[2]), borderWidth: 0 }] }, options: { maintainAspectRatio: false, cutout: '72%', plugins: { legend: { display: false } } } });
    }
    mk('c3', { type: 'bar', data: { labels: ms.map(mLabel), datasets: [
      { label: 'IT energy', data: mrows.map(r => r.it / 1000), backgroundColor: C.purple },
      { label: 'Cooling energy', data: mrows.map(r => r.cool / 1000), backgroundColor: C.blue },
      { label: 'Other', data: mrows.map(r => r.other / 1000), backgroundColor: C.other }] },
      options: opts({ scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true } } }) });
    mk('c4', { data: { labels: ms.map(mLabel), datasets: [
      { type: 'bar', label: 'Cost (NOK M)', data: mrows.map(r => +(r.cost / 1e6).toFixed(2)), backgroundColor: C.purple, yAxisID: 'y' },
      { type: 'line', label: 'Price (NOK/kWh)', data: mrows.map(r => +r.price.toFixed(3)), borderColor: C.amber, backgroundColor: C.amber, yAxisID: 'y1' }] },
      options: opts({ scales: { x: { grid: { display: false } }, y: { beginAtZero: true }, y1: { position: 'right', grid: { display: false }, beginAtZero: true } } }) });
  };

  pages.energy = () => {
    const a = cur(), p = prev(), ms = periodMonths();
    const ren = SITES.reduce((t, s) => t + cur([s]).total * site(s).renewablePct, 0) / a.total;
    app.innerHTML = head('Energy &amp; Cost', 'Consumption and spend across all sites', periodSelect()) + `
      <div class="grid g4">
        ${kpi('Total Energy', gwh(a.total), delta(a.total, p?.total))}
        ${kpi('Total Cost', nokM(a.cost), delta(a.cost, p?.cost, true))}
        ${kpi('Avg. price', `${a.price.toFixed(2)} NOK/kWh`, delta(a.price, p?.price, true))}
        ${kpi('Renewable share', `${ren.toFixed(0)}%`, '<div class="delta">Hydro &amp; wind PPAs</div>')}
      </div>
      <div class="card" style="margin-bottom:12px"><h3>Monthly energy (GWh) and cost (NOK M)</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
      <div class="card"><h3>Cost by site</h3><table><thead><tr><th>Site</th><th>Location</th><th>Energy</th><th>Avg. price</th><th>Est. cost</th></tr></thead><tbody>
        ${SITES.map(s => { const b = cur([s]); return `<tr><td><b>${s}</b></td><td>${site(s).location}</td><td>${gwh(b.total)}</td><td>${b.price.toFixed(2)} NOK/kWh</td><td>NOK ${(b.cost / 1e6).toFixed(2)}M</td></tr>`; }).join('')}
      </tbody></table></div>`;
    const mr = ms.map(m => agg(rows([m])));
    mk('c1', { data: { labels: ms.map(mLabel), datasets: [
      { type: 'line', label: 'Cost (NOK M)', data: mr.map(r => +(r.cost / 1e6).toFixed(2)), borderColor: C.amber, backgroundColor: C.amber, pointRadius: 3 },
      { type: 'bar', label: 'Energy (GWh)', data: mr.map(r => +(r.total / 1000).toFixed(2)), backgroundColor: C.purple }] },
      options: opts({ scales: { x: { grid: { display: false } }, y: { beginAtZero: true } } }) });
  };

  pages.insights = () => {
    const s = state.site, list = anomalies(s), main = D.insights.find(i => i.site === s);
    const done = store.get('actions', {});
    let body;
    if (main) {
      const ex = coolingExcess();
      body = `<div class="banner"><span style="color:${C.amber}">${ic('warn', 26)}</span><div><b>${main.title}</b><p>${main.summary}</p></div><button class="btn" data-action="details">View details ${ic('arrow', 13)}</button></div>
        <div class="grid g21"><div class="card"><h3>Actual vs. expected cooling energy (MWh)</h3><div class="chart tall"><canvas id="c1"></canvas></div></div>
        <div class="card"><h3><span style="color:${C.purple}">${ic('target')}</span> Suggested actions</h3>
          ${main.actions.map((t, i) => `<label class="check ${done[s + i] ? 'done' : ''}"><input type="checkbox" data-check="${s + i}" ${done[s + i] ? 'checked' : ''}><span>${t}</span></label>`).join('')}
          <div class="small" style="margin-top:14px">${main.actions.filter((_, i) => done[s + i]).length} of ${main.actions.length} actions completed</div></div></div>
        <div class="card"><h3><span style="color:${C.green}">${ic('leaf')}</span> Estimated impact</h3><div class="grid g3" style="margin:0">
          <div class="impact"><div class="label">${ic('link', 12)} Potential energy saving</div><div class="value">−${main.savingGWh.toFixed(2)} GWh</div></div>
          <div class="impact"><div class="label">${ic('link', 12)} Estimated cost saving</div><div class="value">NOK ${main.savingNOKM.toFixed(2)}M</div></div>
          <div class="impact"><div class="label">${ic('link', 12)} Estimated emission reduction</div><div class="value">−${main.emissionKt.toFixed(2)} kt CO₂e</div></div></div></div>`;
      setTimeout(() => {
        const d = D.coolingDailyN01, i0 = d.indexOf(ex.win[0]), i1 = d.indexOf(ex.win.at(-1));
        mk('c1', { type: 'line', data: { labels: d.map(x => `${+x.date.slice(8)} Jun`), datasets: [
          { label: 'Actual', data: d.map(x => x.actualMWh), borderColor: C.purple, backgroundColor: C.purple, pointRadius: 0, tension: 0.3, borderWidth: 2 },
          { label: 'Expected (AI model)', data: d.map(x => x.expectedMWh), borderColor: '#cfd2e0', borderDash: [4, 4], pointRadius: 0, tension: 0.4, borderWidth: 1.2 }] },
          options: opts({ scales: { x: { grid: { display: false }, ticks: { maxTicksLimit: 6 } }, y: { min: 150, max: 250 } },
            plugins: { legend: { position: 'bottom' }, overlays: { box: { from: i0, to: i1, label: `+${ex.pct.toFixed(0)}%` } } } }) });
      });
    } else {
      body = `<div class="card">${list.length ? list.map(i => `<div class="insight"><span style="color:${C.amber}">${ic('warn')}</span><div><b>${i.title}</b>${i.summary}</div></div>`).join('') : `<div class="insight"><span style="color:${C.green}">${ic('check')}</span><div><b>No anomalies detected at ${s}</b>Cooling, IT load and PUE are within expected ranges for ${rangeLabel()}.</div></div>`}</div>`;
    }
    app.innerHTML = head('AI Insights', 'AI-powered analysis of energy use, efficiency and capacity', siteSelect() + periodSelect()) + body;
  };

  pages.capacity = () => {
    const inst = D.sites.reduce((t, s) => t + s.installedMW, 0), used = D.sites.reduce((t, s) => t + s.utilisedMW, 0);
    app.innerHTML = head('Capacity', 'Power utilisation and available headroom') + `
      <div class="grid g3">${kpi('Installed capacity', `${inst} MW`, '')}${kpi('Utilised', `${used.toFixed(1)} MW`, `<div class="delta">${(used / inst * 100).toFixed(0)}% of installed</div>`)}${kpi('Available headroom', `${(inst - used).toFixed(1)} MW`, '')}</div>
      <div class="grid g3">${D.sites.map(s => { const u = s.utilisedMW / s.installedMW * 100; return `<div class="card"><h3>${s.id} · ${s.location}</h3>
        <div class="big">${u.toFixed(0)}%</div><div class="small">${s.utilisedMW} of ${s.installedMW} MW</div>
        <div class="bar"><div style="width:${u}%;${u >= 85 ? `background:${C.red}` : ''}"></div></div>
        <div class="small" style="color:#c9ccda">${u >= 85 ? 'Critical – expansion required' : u >= 75 ? 'Approaching limit – plan expansion' : 'Healthy headroom'}</div></div>`; }).join('')}</div>`;
  };

  pages.benchmarking = () => {
    const B = D.benchmark, bulk = B.months.map(m => agg(rows([m])).pue);
    const avg = arr => arr.reduce((a, b) => a + b, 0) / arr.length;
    const bulkAvg = agg(rows(B.months)).pue, mkt = avg(B.pue['Market avg.']);
    const year = agg(rows(B.months)), bulkRen = SITES.reduce((t, s) => t + agg(rows(B.months, [s])).total * site(s).renewablePct, 0) / year.total;
    const p = B.priceN01NOKMWh, hi = p.indexOf(Math.max(...p));
    const winter = (p[hi] + p[(hi + 1) % 12]) / 2, avgP = avg(p);
    let ranking, chart;
    if (state.benchTab === 'efficiency') {
      ranking = [['Bulk (all sites)', bulkAvg], ...Object.entries(B.pue).map(([k, v]) => [k, avg(v)])].map(([k, v], i) => `<div class="rank ${i ? '' : 'me'}"><span>${k}</span><b>${v.toFixed(2)}</b></div>`).join('');
      chart = ['PUE comparison', 'Average PUE (12 months)'];
    } else if (state.benchTab === 'cost') {
      ranking = [['Bulk (all sites)', year.price], ...Object.entries(B.energyCostNOKkWh)].map(([k, v], i) => `<div class="rank ${i ? '' : 'me'}"><span>${k}</span><b>${v.toFixed(2)} NOK/kWh</b></div>`).join('');
      chart = ['Average energy cost (NOK/kWh)', 'Average energy cost (12 months)'];
    } else {
      ranking = [['Bulk (all sites)', bulkRen], ...Object.entries(B.renewablePct)].map(([k, v], i) => `<div class="rank ${i ? '' : 'me'}"><span>${k}</span><b>${v.toFixed(0)}%</b></div>`).join('');
      chart = ['Renewable share (%)', 'Renewable share (12 months)'];
    }
    app.innerHTML = head('Benchmarking', 'Compare Bulk\'s sites with selected competitors and market context', '<select disabled><option>Last 12 months</option></select>') +
      tabs('benchTab', [['efficiency', 'Energy efficiency'], ['cost', 'Energy cost'], ['renewable', 'Renewable share']]) + `
      <div class="grid g21"><div class="card"><h3>${chart[0]}</h3><div class="chart"><canvas id="c1"></canvas></div></div><div class="card"><h3>${chart[1]}</h3>${ranking}</div></div>
      <div class="grid g21"><div class="card"><h3>Electricity price N01 (NOK/MWh)</h3><div class="chart"><canvas id="c2"></canvas></div></div>
      <div class="card"><h3><span style="color:${C.green}">${ic('info')}</span> Key takeaway</h3><p style="font-size:13px;line-height:1.6;color:#c9ccda;margin:0">
        Bulk's average PUE (${bulkAvg.toFixed(2)}) is ${((mkt - bulkAvg) / mkt * 100).toFixed(1)}% better than the market average (${mkt.toFixed(2)}).
        Energy prices were highest in ${mLabel(B.months[hi])}–${mLabel(B.months[(hi + 1) % 12])}, about ${((winter - avgP) / avgP * 100).toFixed(0)}% above the 12-month average price.</p></div></div>`;
    const labels = B.months.map(mLabel);
    if (state.benchTab === 'efficiency') {
      mk('c1', { type: 'line', data: { labels, datasets: [
        { label: 'Bulk (avg.)', data: bulk.map(v => +v.toFixed(3)), borderColor: C.green, backgroundColor: C.green, borderWidth: 2.5, pointRadius: 3 },
        { label: 'Competitor A', data: B.pue['Competitor A'], borderColor: '#e6e7ee', pointRadius: 0, borderWidth: 1 },
        { label: 'Competitor B', data: B.pue['Competitor B'], borderColor: '#e6e7ee', borderDash: [5, 4], pointRadius: 0, borderWidth: 1 },
        { label: 'Market avg.', data: B.pue['Market avg.'], borderColor: '#9aa0b8', pointRadius: 0, borderWidth: 1 }] },
        options: opts({ scales: { x: { grid: { display: false } }, y: { min: 1.15, max: 1.45 } } }) });
    } else {
      const vals = state.benchTab === 'cost' ? [year.price, ...Object.values(B.energyCostNOKkWh)] : [bulkRen, ...Object.values(B.renewablePct)];
      mk('c1', { type: 'bar', data: { labels: ['Bulk', ...Object.keys(B.pue)], datasets: [{ data: vals.map(v => +v.toFixed(2)), backgroundColor: [C.green, '#4a4f6a', '#4a4f6a', '#4a4f6a'], maxBarThickness: 70 }] },
        options: opts({ plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true } } }) });
    }
    mk('c2', { type: 'line', data: { labels, datasets: [{ label: 'NOK/MWh', data: p, borderColor: '#2f9e5f', backgroundColor: 'rgba(47,158,95,0.35)', fill: true, pointRadius: 0, tension: 0, borderWidth: 1.5 }] },
      options: opts({ plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { beginAtZero: true } } }) });
  };

  /* ---------- reports ---------- */
  const allReports = () => [...store.get('reports', []), ...D.reports];
  const fmtDate = d => { const x = new Date(d); return `${x.getDate()} ${MON[x.getMonth()]} ${x.getFullYear()}`; };
  pages.reports = () => {
    app.innerHTML = head('Reports', 'Download generated reports', '<button class="btn" data-action="new-report">New report</button>') +
      `<div class="card">${allReports().map((r, i) => `<div class="report"><span class="ic">${ic('file', 18)}</span><div><b>${r.title}</b><div class="small">${fmtDate(r.date)} · ${r.type}</div></div><button class="btn ghost" data-download="${i}">${ic('down', 13)} Download</button></div>`).join('')}</div>`;
  };

  function reportTable(months) {
    return `<table border="1" cellpadding="6" style="border-collapse:collapse"><tr><th>Site</th><th>Location</th><th>IT (MWh)</th><th>Cooling (MWh)</th><th>Other (MWh)</th><th>Total (MWh)</th><th>PUE</th><th>Cost (NOK)</th></tr>
      ${SITES.map(s => { const a = agg(rows(months, [s])); return `<tr><td>${s}</td><td>${site(s).location}</td><td>${a.it.toFixed(0)}</td><td>${a.cool.toFixed(0)}</td><td>${a.other.toFixed(0)}</td><td>${a.total.toFixed(0)}</td><td>${a.pue.toFixed(2)}</td><td>${a.cost.toLocaleString('en')}</td></tr>`; }).join('')}</table>`;
  }
  function downloadReport(r) {
    if (r.type === 'XLSX' || r.type === 'CSV') {
      const lines = [['Site', 'Location', 'Installed MW', 'Utilised MW', 'Utilisation %', 'Headroom MW'], ...D.sites.map(s => [s.id, s.location, s.installedMW, s.utilisedMW, (s.utilisedMW / s.installedMW * 100).toFixed(1), (s.installedMW - s.utilisedMW).toFixed(1)]),
        [], ['Site', 'Month', 'IT MWh', 'Cooling MWh', 'Other MWh', 'Cost NOK'], ...D.monthly.map(m => [m.site, m.month, m.itMWh, m.coolingMWh, m.otherMWh, m.costNOK])];
      const blob = new Blob([lines.map(l => l.join(',')).join('\n')], { type: 'text/csv' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = r.title.replace(/[^\w]+/g, '_') + '.csv'; a.click();
      return;
    }
    let months = MONTHS.slice(-6), extra = '';
    if (r.kind === 'monthly' && r.month) months = [r.month];
    if (r.kind === 'anomaly') { const i = D.insights[0]; extra = `<h2>${i.title}</h2><p>${i.summary}</p><ul>${i.actions.map(a => `<li>${a}</li>`).join('')}</ul><p>Potential saving: ${i.savingGWh} GWh · NOK ${i.savingNOKM}M · ${i.emissionKt} kt CO₂e</p>`; }
    if (r.kind === 'sustainability') { const a = agg(rows(months)); extra = `<p>Estimated emissions: ${(a.total * D.emissionFactorTPerMWh / 1000).toFixed(2)} kt CO₂e (factor ${D.emissionFactorTPerMWh} t/MWh).</p>`; }
    const site_ = r.site && r.site !== 'All' ? ` – ${r.site}` : '';
    const w = window.open('', '_blank');
    w.document.write(`<html><head><title>${r.title}</title><style>body{font-family:Arial,sans-serif;padding:30px;color:#111}h1{margin-bottom:0}</style></head><body>
      <h1>${r.title}${site_}</h1><p style="color:#666">Bulk Energy Intelligence · generated ${fmtDate(r.date)} · period ${mLong(months[0])}${months.length > 1 ? ' – ' + mLong(months.at(-1)) : ''}</p>
      ${reportTable(months)}${extra}<p style="color:#666;font-size:12px">Use your browser's "Save as PDF" in the print dialog.</p><script>setTimeout(()=>print(),300)<\/script></body></html>`);
    w.document.close();
  }
  function newReportModal() {
    const bg = document.createElement('div'); bg.className = 'modal-bg';
    bg.innerHTML = `<div class="modal"><h3 style="margin-top:0">New report</h3>
      <label>Report type</label><select id="rt"><option value="monthly">Monthly energy report (PDF)</option><option value="sustainability">Sustainability summary (PDF)</option><option value="anomaly">Cooling anomaly analysis (PDF)</option><option value="capacity">Capacity forecast (XLSX/CSV)</option></select>
      <label>Month</label><select id="rm">${MONTHS.slice().reverse().map(m => `<option value="${m}">${mLong(m)}</option>`).join('')}</select>
      <div class="row-end"><button class="btn ghost" data-close>Cancel</button><button class="btn" id="rc">Generate</button></div></div>`;
    document.body.appendChild(bg);
    bg.addEventListener('click', e => { if (e.target === bg || e.target.hasAttribute('data-close')) bg.remove(); });
    bg.querySelector('#rc').onclick = () => {
      const kind = bg.querySelector('#rt').value, month = bg.querySelector('#rm').value;
      const names = { monthly: `Monthly energy report – ${mLong(month).replace(' ', ' ')}`, sustainability: `Sustainability summary – ${mLong(month)}`, anomaly: 'N01 cooling anomaly analysis', capacity: 'Capacity forecast' };
      const r = { title: names[kind], date: new Date().toISOString().slice(0, 10), type: kind === 'capacity' ? 'XLSX' : 'PDF', kind, month };
      store.set('reports', [r, ...store.get('reports', [])]);
      bg.remove(); render();
    };
  }
  function detailsModal() {
    const i = D.insights[0], ex = coolingExcess();
    const bg = document.createElement('div'); bg.className = 'modal-bg';
    bg.innerHTML = `<div class="modal"><h3 style="margin-top:0">${i.title} – ${i.site}</h3>
      <p>${i.summary}</p>
      <p><b>Method:</b> the expected curve is a model of cooling energy given IT load and outdoor temperature. Days where actual cooling exceeds expected by more than 5% are flagged.</p>
      <p><b>Flagged days:</b> ${ex.win.map(d => +d.date.slice(8)).join(', ')} Jun (${ex.win.length} days, +${ex.pct.toFixed(1)}% vs. expected).</p>
      <p><b>Excess energy in window:</b> ${ex.win.reduce((t, d) => t + d.actualMWh - d.expectedMWh, 0).toFixed(0)} MWh.</p>
      <p class="small">AI-generated insight. Validate with site engineers before acting.</p>
      <div class="row-end"><button class="btn" data-close>Close</button></div></div>`;
    document.body.appendChild(bg);
    bg.addEventListener('click', e => { if (e.target === bg || e.target.hasAttribute('data-close')) bg.remove(); });
  }

  pages.settings = () => {
    app.innerHTML = head('Settings', 'Data and preferences') + `
      <div class="card" style="font-size:13px;line-height:1.7"><h3>Data source</h3>
        <p>All figures come from <code>js/data.js</code>: ${D.sites.length} sites, ${MONTHS.length} months (${mLong(MONTHS[0])} – ${mLong(MONTHS.at(-1))}), ${D.coolingDailyN01.length} days of N01 cooling data and competitor benchmarks.</p>
        <p>Saved locally in this browser: ticked AI actions and reports you created.</p>
        <button class="btn ghost" data-action="reset">Reset saved actions and reports</button></div>`;
  };

  /* ---------- router & events ---------- */
  function render() {
    charts.forEach(c => c.destroy()); charts = [];
    const page = (location.hash || '#overview').slice(1);
    const fn = pages[page] || pages.overview;
    document.querySelectorAll('[data-page]').forEach(a => a.classList.toggle('active', a.dataset.page === (pages[page] ? page : 'overview')));
    fn(); window.scrollTo(0, 0);
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
    if (t.dataset.go) location.hash = t.dataset.go;
    if (t.dataset.download) downloadReport(allReports()[+t.dataset.download]);
    const a = t.dataset.action;
    if (a === 'new-report') newReportModal();
    if (a === 'details') detailsModal();
    if (a === 'reset') { store.set('actions', {}); store.set('reports', []); alert('Saved actions and reports cleared.'); }
  });
  document.querySelectorAll('[data-icon]').forEach(el => { el.innerHTML = ic(el.dataset.icon); });
  window.addEventListener('hashchange', render);
  render();
})();
