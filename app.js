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
  positions: [],       // legacy: [{ month, top5, top10, top50, ai }]
  positionsDetail: [], // new schema: [{ month, cluster, engine, avgPos, visibility }]
  socdem: [],          // legacy portrait: [{ segment, gender, age, geo, ... }]
  socdemFunnel: [],    // new schema: [{ gender, age, device, visits, requests, nk, ar }]
  fileName: '',
  filters: { months: [], sources: [], pages: [] },
  ui: { heatMetric: 'ar', rankingEngine: '' }
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

  // Optional auxiliary sheets — looked up by name (case-insensitive).
  const findSheet = re => sheets.find(s => re.test(s.name));
  const positionsSheet = findSheet(/позици|positions?/i);
  const socdemSheet    = findSheet(/socdem|соцдем|соц-дем/i);

  return {
    main: target.json,
    positions:       positionsSheet ? parsePositions(wb.Sheets[positionsSheet.name]) : [],
    positionsDetail: positionsSheet ? parsePositionsDetail(wb.Sheets[positionsSheet.name]) : [],
    socdem:          socdemSheet    ? parseSocdem(wb.Sheets[socdemSheet.name])       : [],
    socdemFunnel:    socdemSheet    ? parseSocdemFunnel(wb.Sheets[socdemSheet.name]) : [],
  };
}

function parsePositions(sheet){
  // First column may be unnamed (month). Read as a 2-D array to be robust to header naming.
  const aoa = XLSX.utils.sheet_to_json(sheet, { header:1, defval:null, raw:true });
  if (!aoa.length) return [];
  const header = aoa[0].map(h => h == null ? '' : String(h).trim());
  const idx = {
    top5:  header.findIndex(h => /топ\s*5\b/i.test(h)),
    top10: header.findIndex(h => /топ\s*10\b/i.test(h)),
    top50: header.findIndex(h => /топ\s*50\b/i.test(h)),
    ai:    header.findIndex(h => /ии|ai|нейросет/i.test(h)),
  };
  // If none of the legacy TOP-N columns is present, this is the new long-form
  // schema (date / cluster / engine / avg pos) — handled by parsePositionsDetail.
  if (idx.top5 < 0 && idx.top10 < 0 && idx.top50 < 0 && idx.ai < 0) return [];
  const out = [];
  for (let i = 1; i < aoa.length; i++){
    const row = aoa[i] || [];
    const month = monthKey(row[0]);
    if (!month) continue;
    out.push({
      month,
      top5:  idx.top5  >= 0 ? toNumber(row[idx.top5])  : null,
      top10: idx.top10 >= 0 ? toNumber(row[idx.top10]) : null,
      top50: idx.top50 >= 0 ? toNumber(row[idx.top50]) : null,
      ai:    idx.ai    >= 0 ? toNumber(row[idx.ai])    : null,
    });
  }
  return out;
}

/* New positions schema: long-form rows with date, cluster, engine, avg position, visibility.
 * Recognised columns:
 *   Дата / Месяц / Date; Кластер / Запрос / Cluster / Query;
 *   Поисковая система / Engine / Source (Яндекс|Google);
 *   Средняя позиция / Avg pos / Position; Видимость / Visibility (% in TOP-10).
 * Returns [] when the sheet doesn't carry these columns; legacy TOP-N counts continue
 * to drive the old "Позиции" widget. */
function parsePositionsDetail(sheet){
  const json = XLSX.utils.sheet_to_json(sheet, { defval:null, raw:true });
  if (!json.length) return [];
  const sample = json[0];
  const keys = Object.keys(sample);
  const has = re => keys.some(k => re.test(String(k)));
  // Require at least the position metric. Cluster/engine are optional but typical.
  if (!has(/средняя\s*позиц|avg.?pos|позиция/i)) return [];

  const out = [];
  for (const r of json){
    const month   = monthKey(pickField(r, /^дата$|^месяц$|date|month/i));
    if (!month) continue;
    const cluster = str(pickField(r, /кластер|запрос|cluster|query|keyword/i));
    const engine  = normEngine(str(pickField(r, /поисков|engine|source|система/i)));
    const avgPos  = toNumber(pickField(r, /средняя\s*позиц|avg.?pos|^позиция$/i));
    let visibility= toNumber(pickField(r, /видимост|visibility/i));
    if (visibility != null && Math.abs(visibility) > 1.5) visibility = visibility/100;
    if (avgPos == null) continue;
    out.push({ month, cluster, engine, avgPos, visibility });
  }
  return out;
}

function normEngine(s){
  if (!s) return '';
  const t = s.toLowerCase();
  if (/яндекс|yandex/.test(t)) return 'Яндекс';
  if (/google|гугл/.test(t)) return 'Google';
  return s;
}

function parseSocdem(sheet){
  const json = XLSX.utils.sheet_to_json(sheet, { defval:null, raw:true });
  if (!json.length) return [];
  // If the sheet uses the new funnel schema (no Сегмент column, but has gender/age + metrics),
  // skip legacy portrait extraction so renderSocdem stays hidden.
  const keys = Object.keys(json[0] || {});
  const hasSegment = keys.some(k => /сегмент|segment/i.test(k));
  if (!hasSegment) return [];
  const out = [];
  for (const r of json){
    const seg = pickField(r, /сегмент|segment/i);
    if (seg == null || String(seg).trim() === '') continue;
    out.push({
      segment:   String(seg).trim(),
      gender:    str(pickField(r, /^пол$|gender/i)),
      age:       str(pickField(r, /возраст|age/i)),
      geo:       str(pickField(r, /географ|geo|регион|город/i)),
      interests: str(pickField(r, /интерес/i)),
      loan:      str(pickField(r, /займ|кредит|loan/i)),
      visits:    str(pickField(r, /визит/i)),
      cr:        str(pickField(r, /конверс|^cr$/i)),
    });
  }
  return out;
}

/* New schema: per-segment funnel metrics with gender × age × device.
 * Recognised columns (case-insensitive, partial match):
 *   Пол / Gender; Возраст / Возрастная группа / Age;
 *   Устройство / Тип устройства / Device;
 *   Визиты / Visits; Заявки / Requests; НК / Новые клиенты;
 *   AR / Approve; CR (опционально).
 * Returns [] when none of the funnel-metric columns is present —
 * in that case the legacy "portrait" view is used instead. */
function parseSocdemFunnel(sheet){
  const json = XLSX.utils.sheet_to_json(sheet, { defval:null, raw:true });
  if (!json.length) return [];
  const sample = json[0];
  const keys = Object.keys(sample);
  const has = re => keys.some(k => re.test(String(k)));
  // Need at least one of the funnel metrics + gender or age or device dimension.
  const hasMetric = has(/^нк$|новые\s*клиент|^ar$|approve|^визит|visits/i);
  const hasDim    = has(/^пол$|gender|возраст|age|устройств|device/i);
  if (!hasMetric || !hasDim) return [];

  const out = [];
  for (const r of json){
    const gender = str(pickField(r, /^пол$|gender/i));
    const age    = normAgeBucket(str(pickField(r, /возрастн|^возраст$|age/i)));
    const device = normDevice(str(pickField(r, /устройств|device/i)));
    if (!gender && !age && !device) continue;
    let visits  = toNumber(pickField(r, /^визит|visits/i));
    let reqs    = toNumber(pickField(r, /^заявк|requests?/i));
    let nk      = toNumber(pickField(r, /^нк$|новые\s*клиент/i));
    let ar      = toNumber(pickField(r, /^ar$|approve/i));
    let cr      = toNumber(pickField(r, /^cr$|конверс/i));
    // Skip rows that have no metrics at all
    if (visits == null && reqs == null && nk == null && ar == null) continue;
    // Percentages may come in as 0..100
    if (ar != null && Math.abs(ar) > 1.5) ar = ar/100;
    if (cr != null && Math.abs(cr) > 1.5) cr = cr/100;
    // Derive AR if missing but we have NK + Заявки
    if (ar == null && nk != null && reqs) ar = nk/reqs;
    if (cr == null && reqs != null && visits) cr = reqs/visits;
    out.push({ gender, age, device, visits, reqs, nk, ar, cr });
  }
  return out;
}

function normAgeBucket(s){
  if (!s) return '';
  const t = s.replace(/\s+/g,'').replace(/[–—−]/g,'-');
  // Map a free-form age expression to one of the standard buckets when possible.
  // Multiple buckets (e.g. "22-34, 35-42") collapse to the first one for grouping.
  const buckets = ['18-24','25-34','35-44','45-54','55+'];
  // Take first numeric token to classify; supports "22-34", "55+", "22-34,35-42", "60+"
  const m = t.match(/(\d+)/);
  if (!m) return s;
  const n = parseInt(m[1],10);
  if (n >= 18 && n <= 24) return '18-24';
  if (n >= 25 && n <= 34) return '25-34';
  if (n >= 35 && n <= 44) return '35-44';
  if (n >= 45 && n <= 54) return '45-54';
  if (n >= 55) return '55+';
  return buckets.includes(t) ? t : s;
}
function normGender(s){
  if (!s) return '';
  const t = s.toLowerCase();
  if (/^муж|^m\b|male/.test(t)) return 'Мужской';
  if (/^жен|^ж\b|^f\b|female/.test(t)) return 'Женский';
  return s;
}
function normDevice(s){
  if (!s) return '';
  const t = s.toLowerCase();
  if (/mobile|моб|смартф|phone/.test(t)) return 'Mobile';
  if (/desktop|десктоп|пк|computer/.test(t)) return 'Desktop';
  if (/tablet|планшет/.test(t)) return 'Tablet';
  return s;
}

function pickField(row, re){
  for (const k of Object.keys(row)){
    if (re.test(String(k))) return row[k];
  }
  return null;
}
function str(v){ return v == null ? '' : String(v).trim(); }

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
      state.positions       = Array.isArray(parsed.positions)       ? parsed.positions       : [];
      state.positionsDetail = Array.isArray(parsed.positionsDetail) ? parsed.positionsDetail : [];
      state.socdem          = Array.isArray(parsed.socdem)          ? parsed.socdem          : [];
      state.socdemFunnel    = Array.isArray(parsed.socdemFunnel)    ? parsed.socdemFunnel    : [];
      state.filters = Object.assign({months:[],sources:[],pages:[]}, parsed.filters||{});
      state.ui      = Object.assign({heatMetric:'ar',rankingEngine:''}, parsed.ui||{});
      return true;
    }
  } catch(e){}
  return false;
}
function clearState(){
  localStorage.removeItem(LS_KEY);
  state = { rows:[], positions:[], positionsDetail:[], socdem:[], socdemFunnel:[],
            fileName:'', filters:{months:[],sources:[],pages:[]},
            ui:{heatMetric:'ar',rankingEngine:''} };
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

/* -------------------- Positions (search rankings) -------------------- */
function renderPositions(){
  destroyChart('positions');
  const section = $('#sectionPositions');
  const data = state.positions || [];
  if (!data.length){ if (section) section.hidden = true; return; }
  if (section) section.hidden = false;

  // Order by calendar months we know about; keep unknown months in original order.
  const ordered = [...data].sort((a,b) => monthSortIdx(a.month) - monthSortIdx(b.month));
  const labels = ordered.map(d => d.month);

  // KPI mini-cards: show last month value and delta vs previous month for each series.
  const last = ordered[ordered.length-1] || {};
  const prev = ordered[ordered.length-2] || {};
  const series = [
    { key:'top5',  label:'ТОП-5',           color:'#16a34a' },
    { key:'top10', label:'ТОП-10',          color:'#2563eb' },
    { key:'top50', label:'ТОП-50',          color:'#7c3aed' },
    { key:'ai',    label:'Упоминания в ИИ', color:'#d97706' },
  ];
  const kpiHost = $('#positionsKpis');
  kpiHost.innerHTML = series.map(s => {
    const cur = last[s.key], pr = prev[s.key];
    const d = (cur != null && pr != null && pr !== 0) ? (cur - pr) / Math.abs(pr) : null;
    const cls = d == null ? 'flat' : d > 0.001 ? 'up' : d < -0.001 ? 'down' : 'flat';
    const arrow = cls==='up'?'▲':cls==='down'?'▼':'▬';
    const dTxt = d == null ? '' : `${arrow} ${fmtDelta(d)}`;
    return `<div class="pos-kpi"><span class="dot" style="background:${s.color}"></span>
      <div class="pos-kpi-body">
        <div class="pos-kpi-label">${esc(s.label)}</div>
        <div class="pos-kpi-value">${fmtInt(cur)}</div>
        <div class="kpi-delta ${cls}">${dTxt}</div>
      </div></div>`;
  }).join('');

  // Two y-axes: left for ТОП-5/10/50 (counts of keywords), right for AI mentions.
  const datasets = [
    ...series.slice(0,3).map(s => ({
      type:'line', label:s.label, data: ordered.map(d => d[s.key]),
      borderColor:s.color, backgroundColor:s.color+'cc',
      borderWidth:2, tension:.3, pointRadius:4, yAxisID:'y'
    })),
    { type:'bar', label:'Упоминания в ИИ',
      data: ordered.map(d => d.ai),
      backgroundColor:'#d97706cc', borderColor:'#d97706', borderWidth:1,
      yAxisID:'y1' }
  ];

  charts.positions = new Chart($('#chartPositions'), {
    data:{ labels, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'bottom' } },
      scales:{
        y:  { position:'left',  beginAtZero:true, ticks:{ callback:v => fmtInt(v) },
              title:{display:true, text:'Кол-во ключевых запросов'} },
        y1: { position:'right', beginAtZero:true, grid:{display:false},
              ticks:{ callback:v => fmtInt(v) },
              title:{display:true, text:'Упоминания в ИИ'} }
      }
    }
  });

  // Detailed table.
  const tbl = $('#positionsTable');
  tbl.innerHTML =
    '<thead><tr><th class="txt">Месяц</th><th>ТОП-5</th><th>ТОП-10</th><th>ТОП-50</th><th>Упоминания в ИИ</th></tr></thead>' +
    '<tbody>' + ordered.map(d => `<tr>
      <td class="txt">${esc(d.month)}</td>
      <td>${fmtInt(d.top5)}</td>
      <td>${fmtInt(d.top10)}</td>
      <td>${fmtInt(d.top50)}</td>
      <td>${fmtInt(d.ai)}</td>
    </tr>`).join('') + '</tbody>';
}

/* -------------------- Socdem (segment portraits) -------------------- */
function renderSocdem(){
  const section = $('#sectionSocdem');
  const data = state.socdem || [];
  if (!data.length){ if (section) section.hidden = true; return; }
  if (section) section.hidden = false;

  // Colour per known segment to keep it visually consistent with the rest of the dashboard.
  const segColors = {
    'Прогретые':       '#16a34a',
    'Присматриваются': '#2563eb',
    'Сыкуны':          '#dc2626',
  };

  const host = $('#socdemCards');
  host.innerHTML = data.map(d => {
    const color = segColors[d.segment] || '#7c3aed';
    const geo = d.geo
      ? d.geo.split(/\s*,\s*/).filter(Boolean).map(g => `<span class="tag">${esc(g)}</span>`).join('')
      : '';
    const interests = d.interests
      ? d.interests.split(/\s*,\s*/).filter(Boolean).map(g => `<span class="tag tag-soft">${esc(g)}</span>`).join('')
      : '';
    return `
      <article class="socdem-card" style="--seg:${color}">
        <header class="socdem-head">
          <div class="socdem-seg">${esc(d.segment)}</div>
          <div class="socdem-gender">${esc(d.gender)}${d.age ? ' · '+esc(d.age) : ''}</div>
        </header>
        <dl class="socdem-grid">
          ${geo ? `<div class="socdem-row"><dt>География</dt><dd class="tags">${geo}</dd></div>` : ''}
          ${interests ? `<div class="socdem-row"><dt>Интересы</dt><dd class="tags">${interests}</dd></div>` : ''}
          ${d.loan ? `<div class="socdem-row"><dt>Типичный займ</dt><dd>${esc(d.loan)}</dd></div>` : ''}
          ${d.visits ? `<div class="socdem-row"><dt>Визитов до заявки</dt><dd>${esc(d.visits)}</dd></div>` : ''}
          ${d.cr ? `<div class="socdem-row"><dt>Конверсия в договор</dt><dd><b>${esc(d.cr)}</b></dd></div>` : ''}
        </dl>
      </article>`;
  }).join('');
}

/* -------------------- Audience (Соц-дем funnel) -------------------- */
const AGE_BUCKETS    = ['18-24','25-34','35-44','45-54','55+'];
const GENDER_BUCKETS = ['Мужской','Женский'];
const DEVICE_COLORS  = { 'Mobile':'#2563eb', 'Desktop':'#16a34a', 'Tablet':'#d97706' };

function aggSocdemBy(rows, keyFn){
  const m = new Map();
  for (const r of rows){
    const k = keyFn(r);
    if (k == null) continue;
    if (!m.has(k)) m.set(k, { visits:0, reqs:0, nk:0 });
    const acc = m.get(k);
    if (r.visits != null) acc.visits += r.visits;
    if (r.reqs   != null) acc.reqs   += r.reqs;
    if (r.nk     != null) acc.nk     += r.nk;
  }
  // Recompute CR/AR from totals (more accurate than averaging per-row rates).
  for (const v of m.values()){
    v.cr = v.visits ? v.reqs/v.visits : null;
    v.ar = v.reqs   ? v.nk  /v.reqs   : null;
  }
  return m;
}

function renderAudience(){
  const section = $('#sectionAudience');
  const rows = (state.socdemFunnel || []).map(r => ({
    ...r,
    gender: normGender(r.gender),
    age:    normAgeBucket(r.age),
    device: normDevice(r.device),
  }));
  if (!rows.length){
    if (section) section.hidden = true;
    destroyChart('deviceShare'); destroyChart('deviceCr');
    return;
  }
  if (section) section.hidden = false;

  // 1) Heatmap by Пол × Возраст
  renderHeatmap(rows);

  // 2) Donut: доля визитов по устройствам
  const byDevice = aggSocdemBy(rows.filter(r => r.device), r => r.device);
  const devices = [...byDevice.keys()];
  destroyChart('deviceShare');
  if (devices.length){
    charts.deviceShare = new Chart($('#chartDeviceShare'), {
      type:'doughnut',
      data:{
        labels: devices,
        datasets:[{
          data: devices.map(d => byDevice.get(d).visits),
          backgroundColor: devices.map(d => DEVICE_COLORS[d] || '#7c3aed'),
          borderWidth:1, borderColor:'#fff'
        }]
      },
      options:{
        responsive:true, maintainAspectRatio:false, cutout:'60%',
        plugins:{
          legend:{position:'bottom'},
          tooltip:{ callbacks:{
            label: ctx => {
              const total = ctx.dataset.data.reduce((s,v)=>s+(+v||0),0) || 1;
              const v = +ctx.raw || 0;
              return `${ctx.label}: ${fmtInt(v)} (${(v/total*100).toFixed(1)}%)`;
            }
          }}
        }
      }
    });
  }

  // 3) Bar: CR / AR по устройствам
  destroyChart('deviceCr');
  if (devices.length){
    charts.deviceCr = new Chart($('#chartDeviceCr'), {
      type:'bar',
      data:{
        labels: devices,
        datasets:[
          { label:'CR (визит → заявка), %', data: devices.map(d => (byDevice.get(d).cr||0)*100),
            backgroundColor:'#93c5fd', borderColor:'#2563eb', borderWidth:1 },
          { label:'AR (заявка → НК), %',     data: devices.map(d => (byDevice.get(d).ar||0)*100),
            backgroundColor:'#86efac', borderColor:'#16a34a', borderWidth:1 },
        ]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        plugins:{
          legend:{position:'bottom'},
          tooltip:{ callbacks:{ label: ctx => `${ctx.dataset.label}: ${(+ctx.raw||0).toFixed(2)}%` } }
        },
        scales:{ y:{ beginAtZero:true, ticks:{ callback:v => v+'%' } } }
      }
    });
  }

  // 4) Detailed table
  renderAudienceTable(rows);
}

function renderHeatmap(rows){
  // Pivot AR / NK over Пол × Возраст. Skip rows with unknown gender or age.
  const buckets = new Map();
  for (const r of rows){
    if (!GENDER_BUCKETS.includes(r.gender) || !AGE_BUCKETS.includes(r.age)) continue;
    const k = r.gender + '||' + r.age;
    if (!buckets.has(k)) buckets.set(k, { visits:0, reqs:0, nk:0 });
    const acc = buckets.get(k);
    if (r.visits != null) acc.visits += r.visits;
    if (r.reqs   != null) acc.reqs   += r.reqs;
    if (r.nk     != null) acc.nk     += r.nk;
  }
  for (const v of buckets.values()){ v.ar = v.reqs ? v.nk/v.reqs : null; }

  const metric = state.ui.heatMetric || 'ar';
  // Determine value range for color scaling
  let maxVal = 0;
  for (const v of buckets.values()){
    const x = metric === 'ar' ? (v.ar || 0) : (v.nk || 0);
    if (x > maxVal) maxVal = x;
  }

  const host = $('#heatmap');
  if (!host) return;
  // Build CSS grid: 1 (label) + N gender columns
  const cols = ['', ...GENDER_BUCKETS];
  host.style.gridTemplateColumns = `120px repeat(${GENDER_BUCKETS.length}, 1fr)`;
  let html = '';
  // header row
  html += cols.map((c,i) => `<div class="heatmap-cell head">${i===0?'':esc(c)}</div>`).join('');
  for (const age of AGE_BUCKETS){
    html += `<div class="heatmap-cell head" style="text-align:left">${esc(age)}</div>`;
    for (const g of GENDER_BUCKETS){
      const cell = buckets.get(g+'||'+age);
      if (!cell || (metric==='ar' && cell.ar==null) || (metric==='nk' && !cell.nk)){
        html += `<div class="heatmap-cell empty"><span class="v">—</span><span class="s">нет данных</span></div>`;
        continue;
      }
      const val   = metric === 'ar' ? cell.ar : cell.nk;
      const ratio = maxVal > 0 ? Math.min(1, val / maxVal) : 0;
      const bg    = heatColor(ratio);
      const fg    = ratio > 0.55 ? '#fff' : '#0f172a';
      const main  = metric === 'ar' ? fmtPct(cell.ar) : fmtInt(cell.nk);
      const sub   = metric === 'ar'
        ? `${fmtInt(cell.nk)} НК · ${fmtInt(cell.reqs)} заявок`
        : `AR ${fmtPct(cell.ar)}`;
      html += `<div class="heatmap-cell" style="background:${bg};color:${fg};border-color:${bg}"
                 title="${esc(g)} · ${esc(age)}">
                 <span class="v">${main}</span>
                 <span class="s">${sub}</span>
               </div>`;
    }
  }
  host.innerHTML = html;
}

// Map ratio 0..1 → color from light blue to deep brand-blue.
function heatColor(r){
  // Linear interpolation between #eff6ff and #1e3a8a in sRGB.
  const a = [239,246,255], b = [30,58,138];
  const c = a.map((x,i) => Math.round(x + (b[i]-x) * r));
  return `rgb(${c[0]},${c[1]},${c[2]})`;
}

function renderAudienceTable(rows){
  const tbl = $('#audienceTable');
  if (!tbl) return;
  // Show every row with at least one metric; sort by AR desc, then NK desc.
  const sorted = [...rows]
    .filter(r => r.ar != null || r.nk != null || r.visits != null)
    .sort((a,b) => (b.ar||0) - (a.ar||0) || (b.nk||0) - (a.nk||0));
  tbl.innerHTML =
    '<thead><tr>' +
      '<th class="txt">Пол</th><th class="txt">Возраст</th><th class="txt">Устройство</th>' +
      '<th>Визиты</th><th>Заявки</th><th>НК</th><th>CR</th><th>AR</th>' +
    '</tr></thead><tbody>' +
    sorted.map(r => `<tr>
      <td class="txt">${esc(r.gender||'—')}</td>
      <td class="txt">${esc(r.age||'—')}</td>
      <td class="txt">${esc(r.device||'—')}</td>
      <td>${fmtInt(r.visits)}</td>
      <td>${fmtInt(r.reqs)}</td>
      <td>${fmtInt(r.nk)}</td>
      <td>${r.cr != null ? fmtPct(r.cr) : '—'}</td>
      <td>${r.ar != null ? fmtPct(r.ar) : '—'}</td>
    </tr>`).join('') +
    '</tbody>';
}

/* -------------------- Ranking (Позиции ↔ НК) -------------------- */
function renderRanking(){
  const section = $('#sectionRanking');
  const detail = state.positionsDetail || [];
  const legacy = state.positions || [];
  // We need either the detailed schema (avg position per month) or, as a graceful
  // fallback, the legacy TOP-N counts to drive the stacked-bucket chart.
  if (!detail.length && !legacy.length){
    if (section) section.hidden = true;
    destroyChart('rankingCorr'); destroyChart('rankingBuckets');
    return;
  }
  if (section) section.hidden = false;

  renderRankingEngineFilter(detail);
  renderRankingCorrelation(detail);
  renderRankingBuckets(detail, legacy);
}

function renderRankingEngineFilter(detail){
  const host = $('#rankingEngineFilter');
  if (!host) return;
  const engines = [...new Set(detail.map(d => d.engine).filter(Boolean))];
  if (engines.length < 2){ host.innerHTML = ''; return; }
  if (!state.ui.rankingEngine || !engines.includes(state.ui.rankingEngine)){
    state.ui.rankingEngine = '';
  }
  const opts = [{ v:'', label:'Все системы' }, ...engines.map(e => ({ v:e, label:e }))];
  host.innerHTML = opts.map(o =>
    `<label class="radio"><input type="radio" name="rankingEngine" value="${esc(o.v)}"
       ${state.ui.rankingEngine === o.v ? 'checked' : ''}> ${esc(o.label)}</label>`
  ).join('');
  host.querySelectorAll('input[name=rankingEngine]').forEach(el => {
    el.addEventListener('change', e => {
      state.ui.rankingEngine = e.target.value;
      saveState();
      renderRankingCorrelation(state.positionsDetail || []);
    });
  });
}

function renderRankingCorrelation(detail){
  destroyChart('rankingCorr');
  const canvas = $('#chartRankingCorr');
  if (!canvas) return;

  const engineFilter = state.ui.rankingEngine || '';
  const filtered = engineFilter ? detail.filter(d => d.engine === engineFilter) : detail;

  // Average position per month across the filtered slice (mean of per-row avgPos).
  const posByMonth = new Map();
  for (const d of filtered){
    if (d.avgPos == null) continue;
    if (!posByMonth.has(d.month)) posByMonth.set(d.month, { sum:0, n:0 });
    const a = posByMonth.get(d.month);
    a.sum += d.avgPos; a.n += 1;
  }
  // НК per month from main data, restricted by current page/source filters.
  const monthly = aggregateByMonth(applyFilters(state.rows));

  // Months present in either source, ordered chronologically.
  const monthsSet = new Set([...monthly.map(m=>m.month), ...posByMonth.keys()]);
  const months = sortMonths([...monthsSet]);
  if (!months.length) return;

  const nkSeries  = months.map(m => {
    const hit = monthly.find(x => x.month === m); return hit ? hit.nk : 0;
  });
  const posSeries = months.map(m => {
    const a = posByMonth.get(m); return a ? a.sum/a.n : null;
  });

  const hasPos = posSeries.some(v => v != null);
  const datasets = [
    { type:'bar', label:'НК (Новые клиенты)', data: nkSeries,
      backgroundColor:'#93c5fd', borderColor:'#2563eb', borderWidth:1, yAxisID:'y' },
  ];
  if (hasPos){
    datasets.push({
      type:'line',
      label:'Средняя позиция' + (engineFilter ? ' · '+engineFilter : ''),
      data: posSeries,
      borderColor:'#dc2626', backgroundColor:'#dc2626',
      borderWidth:2, tension:.3, pointRadius:4, spanGaps:true, yAxisID:'y1'
    });
  }

  charts.rankingCorr = new Chart(canvas, {
    data:{ labels: months, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom'},
        tooltip:{ callbacks:{
          label: ctx => ctx.dataset.yAxisID === 'y1'
            ? `${ctx.dataset.label}: ${(+ctx.raw).toFixed(1)}`
            : `${ctx.dataset.label}: ${fmtInt(ctx.raw)}`
        }}
      },
      scales:{
        y:  { position:'left',  beginAtZero:true, ticks:{ callback:v => fmtInt(v) },
              title:{ display:true, text:'НК (шт.)' } },
        y1: { position:'right', beginAtZero:false, reverse:true,
              suggestedMin:1, grid:{ display:false },
              ticks:{ callback:v => Number(v).toFixed(0) },
              title:{ display:true, text:'Средняя позиция (1 — топ)' },
              display: hasPos }
      }
    }
  });
}

function renderRankingBuckets(detail, legacy){
  destroyChart('rankingBuckets');
  const canvas = $('#chartRankingBuckets');
  if (!canvas) return;

  // Prefer the detailed schema: bucket each query by avg position into TOP-3 / TOP-10 / TOP-30 / 30+.
  // Fall back to legacy TOP5/TOP10/TOP50 monthly counts when only those exist.
  const buckets = ['ТОП-3','ТОП-10','ТОП-30','За ТОП-30'];
  const colors  = { 'ТОП-3':'#16a34a', 'ТОП-10':'#65a30d', 'ТОП-30':'#d97706', 'За ТОП-30':'#94a3b8' };

  let labels, dataByBucket;

  if (detail.length){
    // Group rows by month, then count keywords per bucket.
    const byMonth = new Map();
    for (const d of detail){
      if (d.avgPos == null) continue;
      if (!byMonth.has(d.month)) byMonth.set(d.month, { 'ТОП-3':0,'ТОП-10':0,'ТОП-30':0,'За ТОП-30':0 });
      const acc = byMonth.get(d.month);
      const p = d.avgPos;
      if (p <= 3)       acc['ТОП-3']++;
      else if (p <= 10) acc['ТОП-10']++;
      else if (p <= 30) acc['ТОП-30']++;
      else              acc['За ТОП-30']++;
    }
    labels = sortMonths([...byMonth.keys()]);
    dataByBucket = Object.fromEntries(buckets.map(b => [b, labels.map(m => byMonth.get(m)[b])]));
  } else {
    // Legacy: derive disjoint buckets from cumulative TOP-N counts so the stack reads naturally.
    const ordered = [...legacy].sort((a,b) => monthSortIdx(a.month) - monthSortIdx(b.month));
    labels = ordered.map(d => d.month);
    const t5  = ordered.map(d => +d.top5  || 0);
    const t10 = ordered.map(d => +d.top10 || 0);
    const t50 = ordered.map(d => +d.top50 || 0);
    dataByBucket = {
      'ТОП-3':     t5,                                                     // approximate ТОП-3 ≈ ТОП-5 counts
      'ТОП-10':    t10.map((v,i) => Math.max(0, v - t5[i])),
      'ТОП-30':    t50.map((v,i) => Math.max(0, v - t10[i])),              // counts in ТОП-50 but outside ТОП-10
      'За ТОП-30': labels.map(() => 0),
    };
  }

  const datasets = buckets.map(b => ({
    label: b, data: dataByBucket[b],
    backgroundColor: colors[b], borderColor: colors[b], borderWidth:1, stack:'pos'
  }));

  charts.rankingBuckets = new Chart(canvas, {
    type:'bar',
    data:{ labels, datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom'},
        tooltip:{ callbacks:{ label: ctx => `${ctx.dataset.label}: ${fmtInt(ctx.raw)} запросов` } }
      },
      scales:{
        x:{ stacked:true },
        y:{ stacked:true, beginAtZero:true, ticks:{ callback:v => fmtInt(v) },
            title:{ display:true, text:'Кол-во запросов' } }
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

  // 4. Positions ↔ NK correlation alert (cross-module: Позиции + Воронка)
  // Scans the detailed positions sheet for (cluster, engine) pairs whose avg position
  // improved most month-over-month and ties that to the parallel change in total НК.
  const detail = state.positionsDetail || [];
  if (detail.length && monthly.length >= 2){
    const m1 = monthly[monthly.length-2].month;
    const m2 = monthly[monthly.length-1].month;
    // Key by cluster+engine so a cluster ranked in both Яндекс and Google produces
    // separate comparison candidates and the alert never misattributes the engine.
    const byCe = new Map();
    for (const d of detail){
      if (d.avgPos == null || !d.cluster) continue;
      if (d.month !== m1 && d.month !== m2) continue;
      const k = (d.cluster||'') + '||' + (d.engine||'');
      if (!byCe.has(k)) byCe.set(k, { cluster:d.cluster, engine:d.engine });
      byCe.get(k)[d.month] = d;
    }
    let best = null;
    for (const v of byCe.values()){
      if (!v[m1] || !v[m2]) continue;
      const dPos = v[m1].avgPos - v[m2].avgPos; // positive = поднялись (меньше = лучше)
      if (best == null || dPos > best.dPos){
        best = { name:v.cluster, engine:v.engine, dPos, before:v[m1].avgPos, after:v[m2].avgPos };
      }
    }
    if (best && best.dPos > 0.5){
      const dNk = monthly[monthly.length-1].nk - monthly[monthly.length-2].nk;
      const dV  = monthly[monthly.length-1].visits - monthly[monthly.length-2].visits;
      const trafGrowth = monthly[monthly.length-2].visits ? dV / monthly[monthly.length-2].visits : null;
      const tgTxt = trafGrowth != null ? ` Трафик: ${fmtDelta(trafGrowth)}.` : '';
      const nkTxt = dNk > 0
        ? ` Это совпало с приростом <b>+${fmtInt(dNk)}</b> Новых Клиентов.`
        : ` При этом база НК изменилась на <b>${dNk>=0?'+':''}${fmtInt(dNk)}</b>.`;
      const engineTxt = best.engine ? ` в ${esc(best.engine)}` : '';
      li(ul, `🚀 Мы выросли${engineTxt} по кластеру <b>${esc(best.name)}</b>: средняя позиция улучшилась с ${best.before.toFixed(1)} до ${best.after.toFixed(1)} (${m1} → ${m2}).${tgTxt}${nkTxt}`);
    }
  }

  // 5. Best-AR socdem segment (cross-module: Соц-дем + Воронка)
  const sd = state.socdemFunnel || [];
  if (sd.length){
    const segs = sd
      .map(r => ({ ...r, gender: normGender(r.gender), age: normAgeBucket(r.age), device: normDevice(r.device) }))
      .filter(r => r.ar != null && (r.reqs == null || r.reqs >= 5));
    if (segs.length){
      segs.sort((a,b) => b.ar - a.ar);
      const top = segs[0];
      const label = [top.gender, top.age, top.device].filter(Boolean).join(', ');
      if (label){
        li(ul, `🎯 Сегмент <b>${esc(label)}</b> показывает самый высокий AR — <b>${fmtPct(top.ar)}</b>${top.nk?` (${fmtInt(top.nk)} НК)`:''}. Рекомендуется адаптировать посадочные страницы и креативы под эту аудиторию.`);
      }
    }

    // 6. Mobile funnel warning: high traffic share but low AR vs Desktop.
    const byDev = aggSocdemBy(sd.map(r => ({...r, device: normDevice(r.device)})).filter(r => r.device), r => r.device);
    const mob = byDev.get('Mobile'), desk = byDev.get('Desktop');
    if (mob && desk && mob.visits && desk.visits){
      const totalV = mob.visits + desk.visits + (byDev.get('Tablet')?.visits || 0);
      const mobShare = mob.visits / totalV;
      if (mob.ar != null && desk.ar != null && desk.ar > 0 && mob.ar < desk.ar * 0.7 && mobShare > 0.4){
        const factor = (desk.ar / mob.ar).toFixed(1);
        li(ul, `🚨 Внимание: доля мобильного трафика — <b>${(mobShare*100).toFixed(0)}%</b>, но AR на Mobile (<b>${fmtPct(mob.ar)}</b>) в ${factor}× ниже Desktop (<b>${fmtPct(desk.ar)}</b>). Требуется проверка мобильной воронки и UX лендингов.`);
      }
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
  renderAudience();
  renderRanking();
  renderPositions();
  renderSocdem();
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
    const parsed = parseWorkbook(buf);
    const rows = normalize(parsed.main);
    if (!rows.length) throw new Error('После очистки не осталось строк (проверьте поле "Месяц").');
    state.rows = rows;
    state.positions       = parsed.positions       || [];
    state.positionsDetail = parsed.positionsDetail || [];
    state.socdem          = parsed.socdem          || [];
    state.socdemFunnel    = parsed.socdemFunnel    || [];
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

  // Heatmap metric toggle (Аудитория)
  document.addEventListener('change', e => {
    const t = e.target;
    if (t && t.name === 'heatMetric'){
      state.ui.heatMetric = t.value;
      saveState();
      renderAudience();
    }
  });

  if (loadState()) renderAll();
  else tryLoadDefault();
}
document.addEventListener('DOMContentLoaded', init);

})();
