/* SEO Dashboard — single-file app.
 * Adapts to two data shapes:
 *  A) "Spec" shape: rows with columns Месяц, Источник трафика, Страница/продукт,
 *     Название рекламной кампании, Показы, Клики, CTR, Визиты,
 *     Заявки онлайн-выдача, Заявки офлайн-выдача, CR онлайн, CR офлайн,
 *     НК онлайн, НК офлайн, AR онлайн, AR офлайн.
 *  B) "Aggregate" shape used by bundled datank.xlsx (Metrics sheet) — one row per month
 *     with SEO-трафик всего, Визиты на продуковые страницы, Заявки всего,
 *     Договоры всего, Новые клиенты (договор), and per-segment НК columns.
 * Aggregate rows are expanded into per-segment "products" so the same code path
 * powers all charts.
 */
(() => {
'use strict';

const LS_KEY = 'pril_seo_dashboard_v1';
const MONTH_ORDER = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                     'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
const SOURCE_COLORS = { 'Яндекс':'#fc3f1d', 'Google':'#4285f4', 'Organic':'#10b981',
                        'Директ':'#ffcc00', 'Yandex':'#fc3f1d' };

let state = {
  rows: [],            // normalized rows
  fileName: '',
  filters: { months: [], sources: [], pages: [] }
};
const charts = {};

/* -------------------- Utilities -------------------- */
const $ = sel => document.querySelector(sel);
const fmtInt = n => (n==null||isNaN(n)) ? '—' : new Intl.NumberFormat('ru-RU').format(Math.round(n));
const fmtPct = n => (n==null||isNaN(n)) ? '—' : (n*100).toFixed(2).replace('.',',') + '%';
const fmtDelta = n => (n>0?'+':'') + (n*100).toFixed(1).replace('.',',') + '%';
const esc = s => String(s).replace(/[&<>"']/g, c => (
  { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]
));

function toNumber(v){
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  if (v instanceof Date) return null;
  let s = String(v).trim().replace(/\s+/g,'').replace(/%/g,'').replace(',', '.');
  if (s === '' || s === '-' || s === '—') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function looksLikePct(colName){
  return /CTR|CR|AR|конверс|approve/i.test(colName);
}
function monthKey(m){
  if (m == null) return null;
  const s = String(m).trim();
  if (!s) return null;
  return s;
}
function monthSortIdx(m){
  const i = MONTH_ORDER.indexOf(m);
  return i < 0 ? 99 + (m||'').toString().charCodeAt(0) : i;
}
function sortMonths(arr){
  return [...arr].sort((a,b) => monthSortIdx(a) - monthSortIdx(b));
}
function sumBy(rows, key){
  let s = 0, found = false;
  for (const r of rows){ const v = r[key]; if (v != null && !isNaN(v)) { s += +v; found = true; } }
  return found ? s : null;
}
function safeDiv(a,b){ return (b && !isNaN(b) && !isNaN(a)) ? a/b : null; }

/* -------------------- ETL: parse + normalize -------------------- */
function parseWorkbook(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type:'array', cellDates:true });
  // Pick best sheet: prefer a sheet whose columns look like the spec; else first non-empty.
  const sheets = wb.SheetNames.map(n => ({
    name:n,
    json: XLSX.utils.sheet_to_json(wb.Sheets[n], { defval:null, raw:true })
  })).filter(s => s.json.length > 0);
  if (!sheets.length) throw new Error('В файле нет данных');

  const isSpec = s => s.json[0] && Object.keys(s.json[0]).some(k => /Источник трафика/i.test(k))
                     && Object.keys(s.json[0]).some(k => /НК/i.test(k));
  const target = sheets.find(isSpec) || sheets.find(s => /metrics/i.test(s.name)) || sheets[0];
  return target.json;
}

function normalize(raw){
  // Drop rows with empty Месяц
  const filtered = raw.filter(r => {
    const k = Object.keys(r).find(k => /^месяц$/i.test(String(k).trim()));
    return k && monthKey(r[k]) != null && String(r[k]).trim() !== '';
  });

  // Detect shape
  const cols = filtered.length ? Object.keys(filtered[0]) : [];
  const hasSpec = cols.some(c => /источник трафика/i.test(c));
  return hasSpec ? normalizeSpec(filtered) : normalizeAggregate(filtered);
}

function normalizeSpec(rows){
  return rows.map(r => {
    const out = { _shape:'spec' };
    for (const k of Object.keys(r)){
      const key = k.trim();
      if (/^месяц$/i.test(key))           out['Месяц'] = monthKey(r[k]);
      else if (/источник трафика/i.test(key)) out['Источник'] = r[k] == null ? null : String(r[k]).trim();
      else if (/страница|продукт/i.test(key)) out['Страница'] = r[k] == null ? null : String(r[k]).trim();
      else if (/кампани|кластер|категори/i.test(key)) out['Кластер'] = r[k] == null ? null : String(r[k]).trim();
      else {
        // numeric coercion; percentages stored as fraction 0..1
        let n = toNumber(r[k]);
        if (n != null && looksLikePct(key) && Math.abs(n) > 1.5) n = n/100;
        out[key] = n;
      }
    }
    // derived totals
    out['Заявки'] = (out['Заявки онлайн-выдача']||0) + (out['Заявки офлайн-выдача']||0) || null;
    out['НК']     = (out['НК онлайн']||0) + (out['НК офлайн']||0) || null;
    return out;
  });
}

function normalizeAggregate(rows){
  // The bundled datank.xlsx has one row per month with split per "segment".
  // Expand into per-segment rows so charts/filters work uniformly.
  const segments = [
    { key:'Прогретые',       nk:'Договоры НК Прогретые',       all:'Договоры всего Прогретые' },
    { key:'Присматриваются', nk:'Договоры НК присматривются',  all:'Договоры всего присматривются' },
    { key:'Сыкуны',          nk:'Договоры НК сыкуны',          all:'Договоры всего сыкуны' },
  ];
  const out = [];
  for (const r of rows){
    const month = monthKey(r['Месяц']);
    const visitsTotal  = toNumber(r['SEO-трафик всего']);
    const visitsProd   = toNumber(r['Визиты на продуковые страницы']);
    const reqsTotal    = toNumber(r['Заявки всего']);
    const dealsTotal   = toNumber(r['Договоры всего']);
    const nkTotal      = toNumber(r['Новые клиенты (договор)']);

    // Sums of per-segment fields (used to allocate proportionally).
    const sumNk   = segments.reduce((s,seg)=> s + (toNumber(r[seg.nk])||0), 0);
    const sumAll  = segments.reduce((s,seg)=> s + (toNumber(r[seg.all])||0), 0);

    for (const seg of segments){
      const nk      = toNumber(r[seg.nk]) || 0;
      const dealsAll= toNumber(r[seg.all]) || 0;
      const share   = sumNk ? nk/sumNk : (sumAll ? dealsAll/sumAll : 1/segments.length);
      const visits  = (visitsProd != null ? visitsProd : visitsTotal) * share;
      const requests= reqsTotal != null ? reqsTotal * share : null;
      const deals   = dealsAll;
      out.push({
        _shape:'aggregate',
        'Месяц': month,
        'Источник': 'Organic',           // file doesn't split by engine
        'Страница': seg.key,             // segment is treated as "product"
        'Кластер':  seg.key,
        'Показы':   visitsTotal != null ? visitsTotal * share : null,
        'Клики':    visits,
        'CTR':      visitsTotal ? (visits / (visitsTotal * share)) : null,
        'Визиты':   visits,
        'Заявки онлайн-выдача': requests,
        'Заявки офлайн-выдача': 0,
        'Заявки':   requests,
        'CR онлайн': safeDiv(requests, visits),
        'CR офлайн': null,
        'НК онлайн': nk,
        'НК офлайн': 0,
        'НК':       nk,
        'AR онлайн': safeDiv(nk, requests),
        'AR офлайн': null,
        'Договоры всего': deals,
      });
    }
    // Sanity: if file had nkTotal but segments were missing, fall back to single "Все" row
    if (sumNk === 0 && nkTotal){
      out.push({
        _shape:'aggregate', 'Месяц':month, 'Источник':'Organic', 'Страница':'Все', 'Кластер':'Все',
        'Показы':visitsTotal, 'Клики':visitsProd, 'Визиты':visitsProd, 'CTR':safeDiv(visitsProd,visitsTotal),
        'Заявки':reqsTotal, 'Заявки онлайн-выдача':reqsTotal, 'Заявки офлайн-выдача':0,
        'CR онлайн':safeDiv(reqsTotal, visitsProd), 'CR офлайн':null,
        'НК':nkTotal, 'НК онлайн':nkTotal, 'НК офлайн':0,
        'AR онлайн':safeDiv(nkTotal, reqsTotal), 'AR офлайн':null,
        'Договоры всего':dealsTotal,
      });
    }
  }
  return out;
}

/* -------------------- State persistence -------------------- */
function saveState(){
  try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch(e){ /* quota */ }
}
function loadState(){
  try {
    const s = localStorage.getItem(LS_KEY);
    if (!s) return false;
    const parsed = JSON.parse(s);
    if (parsed && Array.isArray(parsed.rows) && parsed.rows.length){
      state = Object.assign(state, parsed);
      state.filters = Object.assign({months:[],sources:[],pages:[]}, parsed.filters||{});
      return true;
    }
  } catch(e){}
  return false;
}
function clearState(){
  localStorage.removeItem(LS_KEY);
  state = { rows:[], fileName:'', filters:{months:[],sources:[],pages:[]} };
}

/* -------------------- Filters UI -------------------- */
function uniqueValues(key){
  const set = new Set();
  for (const r of state.rows){ if (r[key] != null && r[key] !== '') set.add(r[key]); }
  return [...set];
}
function renderFilters(){
  const months  = sortMonths(uniqueValues('Месяц'));
  const sources = uniqueValues('Источник');
  const pages   = uniqueValues('Страница');

  // default: select all if filter empty
  if (!state.filters.months.length)  state.filters.months  = [...months];
  if (!state.filters.sources.length) state.filters.sources = [...sources];
  if (!state.filters.pages.length)   state.filters.pages   = [...pages];

  // prune stale selections
  state.filters.months  = state.filters.months.filter(m => months.includes(m));
  state.filters.sources = state.filters.sources.filter(s => sources.includes(s));
  state.filters.pages   = state.filters.pages.filter(p => pages.includes(p));

  paintChips($('#filterMonths'),  months,  state.filters.months,  v => toggleFilter('months',v));
  paintChips($('#filterSources'), sources, state.filters.sources, v => toggleFilter('sources',v), v => 'src-'+slug(v));
  paintChips($('#filterPages'),   pages,   state.filters.pages,   v => toggleFilter('pages',v));
}
function slug(v){ return String(v).toLowerCase().replace(/[^a-zа-я0-9]+/gi,''); }
function paintChips(host, values, selected, onClick, extraCls){
  host.innerHTML = '';
  for (const v of values){
    const el = document.createElement('button');
    el.type = 'button';
    el.className = 'chip' + (selected.includes(v) ? ' on' : '') + (extraCls ? ' '+extraCls(v) : '');
    el.textContent = v;
    el.onclick = () => onClick(v);
    host.appendChild(el);
  }
}
function toggleFilter(kind, v){
  const arr = state.filters[kind];
  const i = arr.indexOf(v);
  if (i>=0) arr.splice(i,1); else arr.push(v);
  saveState();
  renderAll();
}
function clearFilters(){
  state.filters = { months:[], sources:[], pages:[] };
  saveState();
  renderAll();
}

/* -------------------- Aggregation -------------------- */
function applyFilters(rows){
  const f = state.filters;
  return rows.filter(r =>
    (!f.months.length  || f.months.includes(r['Месяц'])) &&
    (!f.sources.length || f.sources.includes(r['Источник'])) &&
    (!f.pages.length   || f.pages.includes(r['Страница']))
  );
}

function aggregate(rows){
  const visits   = sumBy(rows,'Визиты') || 0;
  const shows    = sumBy(rows,'Показы') || 0;
  const clicks   = sumBy(rows,'Клики') || visits;
  const reqsOn   = sumBy(rows,'Заявки онлайн-выдача') || 0;
  const reqsOff  = sumBy(rows,'Заявки офлайн-выдача') || 0;
  const reqs     = reqsOn + reqsOff || (sumBy(rows,'Заявки') || 0);
  const nkOn     = sumBy(rows,'НК онлайн') || 0;
  const nkOff    = sumBy(rows,'НК офлайн') || 0;
  const nk       = nkOn + nkOff;
  const cr       = safeDiv(reqs, visits);
  const ar       = safeDiv(nk,   reqs);
  return { visits, shows, clicks, reqs, nk, cr, ar };
}

function aggregateByMonth(rows){
  const m = new Map();
  for (const r of rows){
    if (!m.has(r['Месяц'])) m.set(r['Месяц'], []);
    m.get(r['Месяц']).push(r);
  }
  const out = [];
  for (const month of sortMonths([...m.keys()])){
    out.push(Object.assign({ month }, aggregate(m.get(month))));
  }
  return out;
}

function aggregateBy(rows, key){
  const m = new Map();
  for (const r of rows){
    const k = r[key] || '—';
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  const out = [];
  for (const [k, list] of m.entries()){
    out.push(Object.assign({ key:k }, aggregate(list)));
  }
  return out;
}

function aggregateByMonthSource(rows){
  const m = new Map();
  for (const r of rows){
    const key = r['Месяц'] + '||' + (r['Источник']||'—');
    if (!m.has(key)) m.set(key, { month:r['Месяц'], source:r['Источник']||'—', list:[] });
    m.get(key).list.push(r);
  }
  return [...m.values()].map(g => Object.assign({ month:g.month, source:g.source }, aggregate(g.list)));
}

/* -------------------- KPI cards -------------------- */
function renderKpis(){
  const filtered = applyFilters(state.rows);
  const agg = aggregate(filtered);

  // Compute previous-month delta against the prior month present in unfiltered data,
  // restricted to the same sources/pages.
  const allMonths = sortMonths(uniqueValues('Месяц'));
  const selMonths = sortMonths(state.filters.months.length ? state.filters.months : allMonths);
  const lastSel = selMonths[selMonths.length-1];
  const lastIdx = allMonths.indexOf(lastSel);
  const prevMonth = lastIdx > 0 ? allMonths[lastIdx-1] : null;
  const prevRows = prevMonth ? state.rows.filter(r =>
    r['Месяц'] === prevMonth &&
    (!state.filters.sources.length || state.filters.sources.includes(r['Источник'])) &&
    (!state.filters.pages.length   || state.filters.pages.includes(r['Страница']))
  ) : [];
  const lastRows = state.rows.filter(r =>
    r['Месяц'] === lastSel &&
    (!state.filters.sources.length || state.filters.sources.includes(r['Источник'])) &&
    (!state.filters.pages.length   || state.filters.pages.includes(r['Страница']))
  );
  const prevAgg = prevRows.length ? aggregate(prevRows) : null;
  const lastAgg = lastRows.length ? aggregate(lastRows) : null;

  setKpi('#kpiNk',     fmtInt(agg.nk),      delta(lastAgg?.nk,     prevAgg?.nk));
  setKpi('#kpiVisits', fmtInt(agg.visits),  delta(lastAgg?.visits, prevAgg?.visits));
  setKpi('#kpiCr',     fmtPct(agg.cr),      delta(lastAgg?.cr,     prevAgg?.cr));
  setKpi('#kpiAr',     fmtPct(agg.ar),      delta(lastAgg?.ar,     prevAgg?.ar));
}
function setKpi(sel, value, dlt){
  const card = $(sel);
  card.querySelector('.kpi-value').textContent = value;
  const d = card.querySelector('.kpi-delta');
  if (dlt == null){ d.textContent = ''; d.className='kpi-delta flat'; return; }
  const cls = dlt > 0.001 ? 'up' : dlt < -0.001 ? 'down' : 'flat';
  const arrow = cls==='up'?'▲':cls==='down'?'▼':'▬';
  d.textContent = `${arrow} ${fmtDelta(dlt)} к прошлому месяцу`;
  d.className = 'kpi-delta ' + cls;
}
function delta(curr, prev){
  if (curr == null || prev == null || prev === 0) return null;
  return (curr - prev) / Math.abs(prev);
}

/* -------------------- Charts -------------------- */
function destroyChart(name){ if (charts[name]) { charts[name].destroy(); delete charts[name]; } }

function renderFunnel(){
  destroyChart('funnel');
  const rows = applyFilters(state.rows);
  const a = aggregate(rows);
  const stages = [
    { label:'Показы',  value: a.shows  || a.visits, color:'#bfdbfe' },
    { label:'Визиты',  value: a.visits || a.clicks, color:'#93c5fd' },
    { label:'Заявки',  value: a.reqs,               color:'#60a5fa' },
    { label:'НК',      value: a.nk,                 color:'#2563eb' },
  ].filter(s => s.value && s.value > 0);

  // Use horizontal bar fallback (chartjs-chart-funnel ships funnel chart but
  // a sorted horizontal bar reads identically and avoids extra plugin coupling)
  charts.funnel = new Chart($('#chartFunnel'), {
    type:'bar',
    data:{
      labels: stages.map(s=>s.label),
      datasets:[{
        data: stages.map(s=>s.value),
        backgroundColor: stages.map(s=>s.color),
        borderWidth:0
      }]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{
          label: ctx => {
            const v = ctx.raw, top = stages[0].value;
            const prev = ctx.dataIndex>0 ? stages[ctx.dataIndex-1].value : null;
            const pTop  = top  ? (v/top*100).toFixed(2)+'% от Показов' : '';
            const pPrev = prev ? ' · '+(v/prev*100).toFixed(2)+'% к пред.' : '';
            return `${fmtInt(v)} (${pTop}${pPrev})`;
          }
        }}
      },
      scales:{ x:{ ticks:{ callback:v => fmtInt(v) } } }
    }
  });
}

function renderTrend(){
  destroyChart('trend');
  const rows = applyFilters(state.rows);
  const months = sortMonths([...new Set(rows.map(r=>r['Месяц']))]);
  const sources = [...new Set(rows.map(r=>r['Источник']||'—'))];
  const ms = aggregateByMonthSource(rows);

  const datasets = sources.map(src => ({
    label: src,
    data: months.map(m => {
      const hit = ms.find(x => x.month===m && x.source===src);
      return hit ? hit.nk : 0;
    }),
    borderColor: SOURCE_COLORS[src] || '#7c3aed',
    backgroundColor: (SOURCE_COLORS[src] || '#7c3aed') + 'cc',
    tension:.3, fill:false, borderWidth:2, pointRadius:4
  }));

  charts.trend = new Chart($('#chartTrend'), {
    type:'line',
    data:{ labels: months, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'bottom'} },
      scales:{ y:{ beginAtZero:true, ticks:{ callback:v => fmtInt(v) } } }
    }
  });
}

function renderCombo(){
  destroyChart('combo');
  const rows = applyFilters(state.rows);
  const monthly = aggregateByMonth(rows);
  charts.combo = new Chart($('#chartCombo'), {
    data:{
      labels: monthly.map(x=>x.month),
      datasets:[
        { type:'bar',  label:'Визиты', data: monthly.map(x=>x.visits),
          backgroundColor:'#93c5fd', yAxisID:'y' },
        { type:'line', label:'НК',     data: monthly.map(x=>x.nk),
          borderColor:'#dc2626', backgroundColor:'#dc2626', borderWidth:2,
          tension:.3, pointRadius:4, yAxisID:'y1' }
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'bottom'} },
      scales:{
        y:  { position:'left',  beginAtZero:true, ticks:{ callback:v => fmtInt(v) },
              title:{display:true, text:'Визиты'} },
        y1: { position:'right', beginAtZero:true, grid:{display:false},
              ticks:{ callback:v => fmtInt(v) }, title:{display:true, text:'НК'} }
      }
    }
  });
}

function renderProducts(){
  destroyChart('products');
  const rows = applyFilters(state.rows);
  const byProd = aggregateBy(rows, 'Страница')
    .filter(g => g.nk > 0 || g.reqs > 0)
    .sort((a,b) => b.nk - a.nk)
    .slice(0, 12);

  charts.products = new Chart($('#chartProducts'), {
    data:{
      labels: byProd.map(g=>g.key),
      datasets:[
        { type:'bar', label:'НК', data: byProd.map(g=>g.nk),
          backgroundColor:'#2563eb', yAxisID:'x', xAxisID:'x' },
        { type:'line', label:'AR (Approve Rate)', data: byProd.map(g => (g.ar||0)*100),
          borderColor:'#16a34a', backgroundColor:'#16a34a',
          borderWidth:2, pointRadius:4, tension:.3, xAxisID:'x2' }
      ]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'bottom'} },
      scales:{
        x:  { position:'bottom', beginAtZero:true, ticks:{ callback:v => fmtInt(v) },
              title:{display:true, text:'НК (шт.)'} },
        x2: { position:'top',    beginAtZero:true, grid:{display:false},
              ticks:{ callback:v => v.toFixed(1)+'%' },
              title:{display:true, text:'AR, %'} }
      }
    }
  });
}

/* -------------------- Insights -------------------- */
function renderInsights(){
  const rows = applyFilters(state.rows);
  const ul = $('#insightsList');
  ul.innerHTML = '';
  if (!rows.length){ ul.innerHTML = '<li class="muted">Нет данных для расчёта инсайтов.</li>'; return; }

  // 1. Visits per 1 NK by source
  const bySrc = aggregateBy(rows, 'Источник').filter(g => g.nk > 0);
  bySrc.sort((a,b) => (a.visits/a.nk) - (b.visits/b.nk));
  for (const g of bySrc){
    const ratio = g.visits / g.nk;
    li(ul, `Источник <b>${esc(g.key)}</b>: на 1 нового клиента приходится <b>${fmtInt(ratio)}</b> визитов (${fmtInt(g.visits)} визитов → ${fmtInt(g.nk)} НК).`);
  }
  if (bySrc.length >= 2){
    const best = bySrc[0], worst = bySrc[bySrc.length-1];
    li(ul, `🏆 Самая «дешёвая» по трафику органика — <b>${esc(best.key)}</b> (в ${(worst.visits/worst.nk / (best.visits/best.nk)).toFixed(2)}× эффективнее, чем ${esc(worst.key)}).`);
  }

  // 2. Best AR cluster / campaign
  const dim = state.rows.some(r => r['Кластер']) ? 'Кластер' : 'Страница';
  const byCluster = aggregateBy(rows, dim)
    .filter(g => g.reqs >= 10 && g.ar != null)
    .sort((a,b) => b.ar - a.ar);
  if (byCluster.length){
    const top = byCluster[0];
    li(ul, `🎯 Кластер с наивысшим AR: <b>${esc(top.key)}</b> — <b>${fmtPct(top.ar)}</b> одобрений (${fmtInt(top.nk)} НК из ${fmtInt(top.reqs)} заявок). Рекомендуется приоритизировать SEO-бюджет на эту группу запросов в следующем месяце.`);
    if (byCluster.length > 1){
      const bot = byCluster[byCluster.length-1];
      li(ul, `⚠️ Самый низкий AR — <b>${esc(bot.key)}</b> (${fmtPct(bot.ar)}). Имеет смысл проверить релевантность посадочных страниц или качество семантики.`);
    }
  }

  // 3. MoM growth of NK
  const monthly = aggregateByMonth(rows);
  if (monthly.length >= 2){
    const a = monthly[monthly.length-2], b = monthly[monthly.length-1];
    const d = delta(b.nk, a.nk);
    if (d != null){
      const sign = d>=0 ? 'выросло' : 'снизилось';
      li(ul, `📈 За период ${esc(a.month)} → ${esc(b.month)} количество НК ${sign} на <b>${fmtDelta(d)}</b> (${fmtInt(a.nk)} → ${fmtInt(b.nk)}).`);
    }
  }
}
// Helper for inserting an insight as <li>. Callers MUST pass HTML-safe content
// (use esc() on any user/data-derived strings). Used only by renderInsights().
function li(ul, html){ const el = document.createElement('li'); el.innerHTML = html; ul.appendChild(el); }

/* -------------------- Data grid -------------------- */
let gridSort = { key:'НК онлайн', dir:'desc' };

function renderGrid(){
  const rows = applyFilters(state.rows);
  const cols = [
    { key:'Месяц', label:'Месяц', txt:true },
    { key:'Источник', label:'Источник', txt:true },
    { key:'Страница', label:'Страница/продукт', txt:true },
    { key:'Кластер',  label:'Кластер',  txt:true },
    { key:'Показы',   label:'Показы' },
    { key:'Визиты',   label:'Визиты' },
    { key:'CTR',      label:'CTR', pct:true },
    { key:'Заявки',   label:'Заявки' },
    { key:'CR онлайн',label:'CR онлайн', pct:true },
    { key:'НК онлайн',label:'НК онлайн' },
    { key:'НК офлайн',label:'НК офлайн' },
    { key:'AR онлайн',label:'AR онлайн', pct:true },
  ];

  // Sort
  const sorted = [...rows].sort((a,b) => {
    const va = a[gridSort.key], vb = b[gridSort.key];
    if (va == null && vb == null) return 0;
    if (va == null) return 1;
    if (vb == null) return -1;
    if (typeof va === 'number' && typeof vb === 'number'){
      return gridSort.dir==='asc' ? va-vb : vb-va;
    }
    return gridSort.dir==='asc'
      ? String(va).localeCompare(String(vb),'ru')
      : String(vb).localeCompare(String(va),'ru');
  });

  const thead = $('#dataGrid thead'), tbody = $('#dataGrid tbody');
  thead.innerHTML = '<tr>' + cols.map(c => {
    const arr = gridSort.key===c.key ? (gridSort.dir==='asc'?'▲':'▼') : '';
    return `<th class="${c.txt?'txt':''}" data-key="${esc(c.key)}">${esc(c.label)} <span class="arr">${arr}</span></th>`;
  }).join('') + '</tr>';
  thead.querySelectorAll('th').forEach(th => th.onclick = () => {
    const k = th.dataset.key;
    if (gridSort.key === k) gridSort.dir = gridSort.dir==='asc'?'desc':'asc';
    else { gridSort.key = k; gridSort.dir = 'desc'; }
    renderGrid();
  });

  tbody.innerHTML = sorted.map(r => '<tr>' + cols.map(c => {
    const v = r[c.key];
    if (c.txt) return `<td class="txt">${v == null ? '' : esc(v)}</td>`;
    if (c.pct) return `<td>${v == null ? '' : fmtPct(v)}</td>`;
    return `<td>${v == null ? '' : fmtInt(v)}</td>`;
  }).join('') + '</tr>').join('');

  $('#gridMeta').textContent = `${sorted.length} строк (после фильтрации)`;
}

/* -------------------- Export -------------------- */
function exportData(){
  const rows = applyFilters(state.rows);
  if (!rows.length){ alert('Нет данных для экспорта'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const out = {};
    for (const k of Object.keys(r)){ if (!k.startsWith('_')) out[k] = r[k]; }
    return out;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'SEO');
  XLSX.writeFile(wb, 'seo_dashboard_export.xlsx');
}

/* -------------------- Top-level render -------------------- */
function renderAll(){
  if (!state.rows.length){
    $('#dashboard').hidden = true;
    $('#dropzone').hidden = false;
    return;
  }
  $('#dashboard').hidden = false;
  $('#dropzone').hidden = true;
  renderFilters();
  renderKpis();
  renderFunnel();
  renderTrend();
  renderCombo();
  renderProducts();
  renderInsights();
  renderGrid();
  $('#srcName').textContent = state.fileName || '—';
  $('#srcRows').textContent = state.rows.length;
}

/* -------------------- File handling -------------------- */
async function loadFile(file){
  const buf = await file.arrayBuffer();
  ingestArrayBuffer(buf, file.name);
}
function ingestArrayBuffer(buf, name){
  try {
    const raw = parseWorkbook(buf);
    const rows = normalize(raw);
    if (!rows.length) throw new Error('После очистки не осталось строк (проверьте поле "Месяц").');
    state.rows = rows;
    state.fileName = name;
    state.filters = { months:[], sources:[], pages:[] };
    saveState();
    renderAll();
  } catch (e){
    alert('Не удалось обработать файл: ' + e.message);
    console.error(e);
  }
}

async function tryLoadDefault(){
  // Try to fetch the bundled datank.xlsx as the default reference dataset.
  try {
    const res = await fetch('datank.xlsx');
    if (!res.ok) return;
    const buf = await res.arrayBuffer();
    ingestArrayBuffer(buf, 'datank.xlsx');
  } catch(e){ /* opened via file:// — user must drop a file */ }
}

/* -------------------- Wire-up -------------------- */
function init(){
  const dz = $('#dropzone');
  ['dragenter','dragover'].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); dz.classList.add('drag');
  }));
  ['dragleave','drop'].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); dz.classList.remove('drag');
  }));
  dz.addEventListener('drop', e => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) loadFile(f);
  });
  $('#fileInput').addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
    if (f) loadFile(f);
  });
  $('#btnReset').addEventListener('click', () => {
    if (!confirm('Сбросить загруженные данные и фильтры?')) return;
    clearState();
    for (const k of Object.keys(charts)) destroyChart(k);
    renderAll();
  });
  $('#btnExport').addEventListener('click', exportData);
  $('#btnClearFilters').addEventListener('click', clearFilters);

  if (loadState()) renderAll();
  else tryLoadDefault();
}
document.addEventListener('DOMContentLoaded', init);

})();
