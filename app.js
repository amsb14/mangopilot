'use strict';

// ── EMBEDDED DATA (519 companies) ─────────────────────────────────────────────
let DB_RAW = { annual: {}, quarterly: {}, stock: {} };
let FILE_LOADED = false;

// ── XLSX PARSER (uses SheetJS loaded from CDN) ───────────────────────────────
async function parseFinancialXLSX(arrayBuffer) {
  const XLSX = window.XLSX;
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

  const sheetNames = wb.SheetNames;
  const getSheet = (keyword) => {
    const name = sheetNames.find(n => n.toLowerCase().includes(keyword.toLowerCase()));
    return name ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null }) : [];
  };

  const incomeAll  = getSheet('Income');
  const cashAll    = getSheet('Cash');
  const balanceAll = getSheet('Balance');
  const stockAll   = getSheet('Stock');
  const dividAll   = getSheet('Dividend');

  const s = v => { const n = parseFloat(v); return (v == null || isNaN(n) || !isFinite(n)) ? null : n; };
  const dateYear = v => { if (!v) return null; if (v instanceof Date) return v.getFullYear(); const d = new Date(v); return isNaN(d.getTime()) ? null : d.getFullYear(); };
  const dateStr  = v => { if (!v) return null; if (v instanceof Date) return v.toISOString().slice(0,10); const d = new Date(v); return isNaN(d.getTime()) ? null : d.toISOString().slice(0,10); };

  // ── Index by [Ticker][year/date] for merging ──
  function indexBy(rows, type, keyFn) {
    const idx = {};
    rows.forEach(r => {
      if ((r.Statement_Type || '').toLowerCase() !== type.toLowerCase()) return;
      const t = r.Ticker; const k = keyFn(r);
      if (!t || !k) return;
      if (!idx[t]) idx[t] = {};
      idx[t][k] = r;
    });
    return idx;
  }

  const incomeAnnual  = indexBy(incomeAll,  'Annual',    r => dateYear(r.Fiscal_Date));
  const cashAnnual    = indexBy(cashAll,     'Annual',    r => dateYear(r.Fiscal_Date));
  const balanceAnnual = indexBy(balanceAll,  'Annual',    r => dateYear(r.Fiscal_Date));
  const incomeQ       = indexBy(incomeAll,   'Quarterly', r => dateStr(r.Fiscal_Date));
  const cashQ         = indexBy(cashAll,     'Quarterly', r => dateStr(r.Fiscal_Date));

  // ── Build annual data ──
  const annual = {};
  const allTickers = new Set([
    ...Object.keys(incomeAnnual),
    ...Object.keys(cashAnnual),
    ...Object.keys(balanceAnnual)
  ]);

  allTickers.forEach(ticker => {
    const incYears = incomeAnnual[ticker]  || {};
    const cfYears  = cashAnnual[ticker]    || {};
    const bsYears  = balanceAnnual[ticker] || {};
    const allYears = new Set([...Object.keys(incYears), ...Object.keys(cfYears), ...Object.keys(bsYears)]);

    const rows = [];
    allYears.forEach(yr => {
      const y   = parseInt(yr);
      const inc = incYears[yr] || {};
      const cf  = cfYears[yr]  || {};
      const bs  = bsYears[yr]  || {};

      const ta = s(bs.TotalAssets);
      const tl = s(bs.TotalLiabilitiesNetMinorityInterest);
      const eq = s(bs.CommonStockEquity) || s(bs.StockholdersEquity) || ((ta && tl) ? ta - tl : null);
      const ocf   = s(cf.OperatingCashFlow);
      const capex = s(cf.CapitalExpenditure);
      let   fcf   = s(cf.FreeCashFlow);
      if (fcf == null && ocf != null && capex != null) fcf = ocf + capex;

      const entry = {
        year: y, ticker,
        revenue:              s(inc.TotalRevenue),
        net_income:           s(inc.NetIncome),
        operating_income:     s(inc.OperatingIncome),
        gross_profit:         s(inc.GrossProfit),
        ebitda:               s(inc.EBITDA),
        ebit:                 s(inc.EBIT),
        eps_diluted:          s(inc.DilutedEPS),
        total_assets:         ta,
        total_liabilities:    tl,
        equity:               eq,
        total_debt:           s(bs.TotalDebt),
        operating_cash_flow:  ocf,
        capital_expenditures: capex,
        free_cash_flow:       fcf,
        interest_expense:     s(inc.InterestExpense),
        tax_provision:        s(inc.TaxProvision),
        shares_diluted:       s(inc.DilutedAverageShares),
      };

      if (entry.revenue != null || entry.net_income != null) rows.push(entry);
    });

    if (rows.length) annual[ticker] = rows.sort((a,b) => a.year - b.year);
  });

  // ── Build quarterly data ──
  const quarterly = {};
  const qTickers = new Set([...Object.keys(incomeQ), ...Object.keys(cashQ)]);
  qTickers.forEach(ticker => {
    const incDates = incomeQ[ticker] || {};
    const cfDates  = cashQ[ticker]   || {};
    const allDates = new Set([...Object.keys(incDates), ...Object.keys(cfDates)]);
    const qRows = [];
    allDates.forEach(dt => {
      const inc = incDates[dt] || {};
      const cf  = cfDates[dt]  || {};
      const entry = {
        date: dt,
        revenue:          s(inc.TotalRevenue),
        net_income:       s(inc.NetIncome),
        operating_income: s(inc.OperatingIncome),
        gross_profit:     s(inc.GrossProfit),
        eps_diluted:      s(inc.DilutedEPS),
        operating_cash_flow: s(cf.OperatingCashFlow),
        free_cash_flow:      s(cf.FreeCashFlow),
      };
      // Keep only entries with at least some data
      if (Object.values(entry).some((v,i) => i > 0 && v != null)) qRows.push(entry);
    });
    if (qRows.length) quarterly[ticker] = qRows.sort((a,b) => b.date.localeCompare(a.date)).slice(0,8);
  });

  // ── Build stock info ──
  const stock = {};
  stockAll.forEach(r => {
    const t = r.Ticker;
    if (!t) return;
    stock[t] = {
      price:            s(r.currentPrice),
      marketCap:        s(r.marketCap),
      pe:               s(r.trailingPE),
      sector:           r.sector || null,
      industry:         r.industry || null,
      currentRatio:     s(r.currentRatio),
      totalDebt:        s(r.totalDebt),
      roe:              s(r.returnOnEquity),
      roa:              s(r.returnOnAssets),
      bookValue:        s(r.bookValue),
      recommendation:   r.recommendationKey || null,
      fiftyTwoWeekLow:  s(r.fiftyTwoWeekLow),
      fiftyTwoWeekHigh: s(r.fiftyTwoWeekHigh),
      dividendYield:    s(r.dividendYield),
    };
  });

  return { annual, quarterly, stock };
}

// ── STATE ─────────────────────────────────────────────────────────────────────
let lang = 'en';
let theme = 'light'; // 'dark' | 'light' — default to light
document.body.classList.add('light');
document.getElementById('themeBtn').textContent = '🌙';
let activeTicker = null;
let cmpList = [];       // tickers in compare bar
let charts = {};
let currentDash = null;
let activeSector = null;
let sidebarQuery = '';

let TICKERS = [];
let STOCK   = {};
let ANNUAL  = {};
let QUARTERLY = {};
let ESTIMATES = {};

function refreshGlobals() {
  TICKERS = Object.keys(DB_RAW.annual).sort();
  STOCK = DB_RAW.stock || {};
  ANNUAL = DB_RAW.annual || {};
  QUARTERLY = DB_RAW.quarterly || {};
  ESTIMATES = DB_RAW.estimates || {};
}

// ── I18N ──────────────────────────────────────────────────────────────────────
const i18n = {
en:{
  sidebarLabel:'Browse Companies', sidebarTickersLabel:'Companies',
  eyebrowTxt:'519 Companies · AI-Powered Analysis',
  welcomeH1:'Financial Intelligence<br>at Your Fingertips',
  welcomeP:'Search for any ticker or browse by sector. Get instant analysis, year-over-year trends, risk scores, and AI-generated insights.',
  statLbl1:'Companies', statLbl2:'Annual Data', statLbl3:'Quarterly Data', statLbl4:'Sectors',
  hw1T:'Search Any Ticker', hw1X:'Type a ticker symbol or company name in the search bar to load its financial dashboard.',
  hw2T:'Compare Companies', hw2X:'Click "+ Compare" on any dashboard to add it to a comparison. Up to 5 companies.',
  hw3T:'Upload Your Own', hw3X:'Use the upload button to add your own CSV or Excel statements with auto schema detection.',
  hdrSearchPlaceholder:'Search ticker or company…',
  sidebarSearchPlaceholder:'Search…',
  langLbl:'العربية',
  themeLight:'☀️', themeDark:'🌙',
  cmpBarLabel:'Comparing:', cmpGoBtn:'Compare →',
  overallScore:'Overall', growth:'Growth', profitability:'Profitability', health:'Health', cashflow:'Cash Flow',
  revenue:'Revenue', netIncome:'Net Income', totalAssets:'Total Assets', equity:'Equity',
  ocf:'Op. Cash Flow', fcf:'Free Cash Flow', opIncome:'Op. Income', ebitda:'EBITDA',
  eps:'EPS (Diluted)', totalDebt:'Total Debt',
  latestYr:'Latest year',
  aiTag:'✦ AI ANALYSIS', aiRegen:'Regenerate',
  revChart:'Revenue & Net Income', cfChart:'Cash Flow Trend',
  yoyChart:'YoY Revenue Growth (%)', radarChart:'Score Breakdown',
  yoyTable:'Year-over-Year Analysis', profTable:'Profitability Metrics', qTable:'Quarterly Data',
  cmpTitle:'Company Comparison', cmpScores:'Scores',
  yr:'Year', revGrowth:'Rev. Growth', niGrowth:'NI Growth', ocfGrowth:'OCF Growth',
  netMargin:'Net Margin', opMargin:'Op. Margin', roe:'ROE', roa:'ROA', dte:'D/E',
  price:'Price', mktCap:'Mkt Cap', pe:'P/E', sector:'Sector', industry:'Industry',
  recommendation:'Analyst', currentRatio:'Curr. Ratio',
  exportCSV:'⬇ CSV', exportPDF:'🖨 PDF', addCmp:'+ Compare', addedCmp:'✓ Added',
  loaderTxt:'Loading analysis…',
  yearsData:'yrs data', date:'Date', noData:'—',
  allSectors:'All Sectors',
},
ar:{
  sidebarLabel:'تصفح الشركات', sidebarTickersLabel:'الشركات',
  eyebrowTxt:'519 شركة · تحليل بالذكاء الاصطناعي',
  welcomeH1:'ذكاء مالي<br>في متناول يدك',
  welcomeP:'ابحث عن أي رمز سهم أو تصفح حسب القطاع. احصل على تحليل فوري واتجاهات سنوية ودرجات مخاطر ورؤى ذكية.',
  statLbl1:'شركة', statLbl2:'بيانات سنوية', statLbl3:'بيانات فصلية', statLbl4:'قطاعات',
  hw1T:'ابحث عن أي رمز', hw1X:'اكتب رمز السهم أو اسم الشركة في شريط البحث لتحميل لوحة التحليل المالي.',
  hw2T:'قارن الشركات', hw2X:'انقر على "+ مقارنة" في أي لوحة لإضافتها. يمكنك مقارنة حتى 5 شركات.',
  hw3T:'ارفع ملفك', hw3X:'استخدم زر الرفع لإضافة قوائمك المالية CSV أو Excel مع اكتشاف تلقائي للمخطط.',
  hdrSearchPlaceholder:'ابحث عن رمز أو شركة…',
  sidebarSearchPlaceholder:'بحث…',
  langLbl:'English',
  themeLight:'☀️', themeDark:'🌙',
  cmpBarLabel:'المقارنة:', cmpGoBtn:'قارن →',
  overallScore:'الإجمالية', growth:'النمو', profitability:'الربحية', health:'الصحة', cashflow:'التدفق',
  revenue:'الإيرادات', netIncome:'صافي الدخل', totalAssets:'إجمالي الأصول', equity:'حقوق الملكية',
  ocf:'التدفق التشغيلي', fcf:'التدفق الحر', opIncome:'الدخل التشغيلي', ebitda:'EBITDA',
  eps:'ربحية السهم', totalDebt:'إجمالي الديون',
  latestYr:'آخر سنة',
  aiTag:'✦ تحليل ذكي', aiRegen:'إعادة توليد',
  revChart:'الإيرادات وصافي الدخل', cfChart:'اتجاه التدفق النقدي',
  yoyChart:'نمو الإيرادات السنوي (%)', radarChart:'تحليل الدرجات',
  yoyTable:'التحليل السنوي المقارن', profTable:'مقاييس الربحية', qTable:'البيانات الفصلية',
  cmpTitle:'مقارنة الشركات', cmpScores:'الدرجات',
  yr:'السنة', revGrowth:'نمو الإيرادات', niGrowth:'نمو الدخل', ocfGrowth:'نمو التدفق',
  netMargin:'هامش صافي', opMargin:'هامش تشغيلي', roe:'العائد على الملكية', roa:'العائد على الأصول', dte:'الدين/الملكية',
  price:'السعر', mktCap:'القيمة السوقية', pe:'مضاعف الربح', sector:'القطاع', industry:'الصناعة',
  recommendation:'توصية', currentRatio:'نسبة التداول',
  exportCSV:'⬇ CSV', exportPDF:'🖨 PDF', addCmp:'+ مقارنة', addedCmp:'✓ أُضيف',
  loaderTxt:'جارٍ تحميل التحليل…',
  yearsData:'سنوات', date:'التاريخ', noData:'—',
  allSectors:'جميع القطاعات',
}};
const T = k => (i18n[lang]||i18n.en)[k] ?? i18n.en[k] ?? k;

// ── THEME ─────────────────────────────────────────────────────────────────────
function toggleTheme() {
  theme = theme === 'dark' ? 'light' : 'dark';
  document.body.classList.toggle('light', theme === 'light');
  document.getElementById('themeBtn').textContent = theme === 'light' ? '🌙' : '☀️';
  // Redraw charts with new colors
  if (currentDash) setTimeout(() => drawCharts(currentDash.ticker, currentDash.rows, currentDash.m), 50);
}

// ── FORMATTING ────────────────────────────────────────────────────────────────
function fmtAI(t) {
  if (!t) return '';
  // Escape HTML
  let s = t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // Block-level: process line by line
  const lines = s.split('\n');
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const trimmed = lines[i].trim();

    // Table: pipe header + separator + body rows
    if (/^\|.+\|$/.test(trimmed) && i + 1 < lines.length && /^\|[\s:|-]+\|$/.test(lines[i+1].trim())) {
      const ths = trimmed.split('|').slice(1, -1).map(c => `<th>${c.trim()}</th>`).join('');
      const rows = []; i += 2;
      while (i < lines.length && /^\|.+\|$/.test(lines[i].trim())) {
        const tds = lines[i].trim().split('|').slice(1, -1).map(c => `<td>${c.trim()}</td>`).join('');
        rows.push(`<tr>${tds}</tr>`); i++;
      }
      out.push(`<table class="ai-md-table"><thead><tr>${ths}</tr></thead><tbody>${rows.join('')}</tbody></table>`);
      continue;
    }

    // Headers — most hashes first; ### and above all use h3 styling (main headings),
    // #### and deeper use smaller h4. Without #### handling, COMPOSER_PROMPT output infinite-loops.
    if (/^#{4,6}\s+/.test(trimmed)) { out.push(`<div class="ai-md-h4">${trimmed.replace(/^#{4,6}\s+/, '')}</div>`); i++; continue; }
    if (/^#{1,3}\s+/.test(trimmed)) { out.push(`<div class="ai-md-h3">${trimmed.replace(/^#{1,3}\s+/, '')}</div>`); i++; continue; }

    // Horizontal rule
    if (/^-{3,}$/.test(trimmed)) { out.push('<hr class="ai-md-hr">'); i++; continue; }

    // Unordered list
    if (/^[-*]\s+/.test(trimmed)) {
      const items = [];
      while (i < lines.length && /^[-*]\s+/.test(lines[i].trim())) {
        items.push(`<li>${lines[i].trim().replace(/^[-*]\s+/, '')}</li>`); i++;
      }
      out.push(`<ul class="ai-md-ul">${items.join('')}</ul>`); continue;
    }

    // Ordered list
    if (/^\d+\.\s+/.test(trimmed)) {
      const items = [];
      while (i < lines.length && /^\d+\.\s+/.test(lines[i].trim())) {
        items.push(`<li>${lines[i].trim().replace(/^\d+\.\s+/, '')}</li>`); i++;
      }
      out.push(`<ol class="ai-md-ol">${items.join('')}</ol>`); continue;
    }

    // Blank line — skip
    if (trimmed === '') { i++; continue; }

    // Prose paragraph: accumulate consecutive non-block lines
    const startI = i;
    const proseLines = [];
    while (i < lines.length) {
      const tt = lines[i].trim();
      if (tt === '' || /^#{1,6}\s/.test(tt) || /^[-*]\s/.test(tt) || /^\d+\.\s/.test(tt) || /^\|.+\|$/.test(tt) || /^-{3,}$/.test(tt)) break;
      proseLines.push(tt); i++;
    }
    if (proseLines.length) out.push(`<p class="ai-md-p">${proseLines.join('<br>')}</p>`);
    if (i === startI) i++; // safety: never let a line loop forever
  }

  let result = out.join('');

  // Inline replacements (run on the assembled output)
  result = result.replace(/\*\*(.+?)\*\*/g, '<strong style="color:var(--accent)">$1</strong>');
  result = result.replace(/(?<!\*)\*([^*\n]+?)\*(?!\*)/g, '<em>$1</em>');
  result = result.replace(/`([^`\n]+)`/g, '<code class="ai-md-code">$1</code>');

  return result;
}

function fM(n) {
  if (n == null || isNaN(n)) return T('noData');
  const a = Math.abs(n), neg = n < 0;
  let s;
  if (a >= 1e12) s = '$' + (a/1e12).toFixed(2) + 'T';
  else if (a >= 1e9) s = '$' + (a/1e9).toFixed(2) + 'B';
  else if (a >= 1e6) s = '$' + (a/1e6).toFixed(1) + 'M';
  else s = '$' + a.toLocaleString(undefined, {maximumFractionDigits:0});
  return neg ? '(' + s + ')' : s;
}
function fP(n, dp=1) { if (n==null||isNaN(n)) return T('noData'); return (n>0?'+':'') + n.toFixed(dp) + '%' }
function fR(n) { if (n==null||isNaN(n)) return T('noData'); return n.toFixed(2) + 'x' }
function fN(n, dp=2) { if (n==null||isNaN(n)) return T('noData'); return n.toFixed(dp) }
function sCol(s) { if(s>=8) return '#10b981'; if(s>=6.5) return '#3b82f6'; if(s>=5) return '#f59e0b'; return '#ef4444'; }
function pc(v) { if(v==null||isNaN(v)) return 'neu'; return v>0?'pos':v<0?'neg':'neu'; }

// ── SIDEBAR ───────────────────────────────────────────────────────────────────
function buildSidebar() {
  if (!TICKERS.length) {
    document.getElementById('sectorList').innerHTML = `<div style="padding:12px 4px;font-size:13px;color:var(--text2)">${lang==='ar'?'ارفع الملف للبدء':'Upload a file to start'}</div>`;
    document.getElementById('sidebarTickers').innerHTML = '';
    return;
  }
  // Sectors
  const sectors = {};
  TICKERS.forEach(t => {
    const sec = STOCK[t]?.sector || 'Other';
    sectors[sec] = (sectors[sec]||0) + 1;
  });

  const sectorEl = document.getElementById('sectorList');
  const allItem = `<div class="sector-item active" id="sector-ALL" onclick="selectSector(null)">${T('allSectors')} <span class="sector-count">${TICKERS.length}</span></div>`;
  const items = Object.entries(sectors)
    .sort((a,b) => b[1]-a[1])
    .map(([sec, cnt]) =>
      `<div class="sector-item" id="sector-${sec.replace(/\s/g,'_')}" onclick="selectSector('${sec}')">${sec} <span class="sector-count">${cnt}</span></div>`
    ).join('');
  sectorEl.innerHTML = allItem + items;

  renderSidebarTickers();
}

function selectSector(sec) {
  activeSector = sec;
  activeTicker = null;
  document.querySelectorAll('.sector-item').forEach(el => el.classList.remove('active'));
  const id = sec ? 'sector-' + sec.replace(/\s/g,'_') : 'sector-ALL';
  document.getElementById(id)?.classList.add('active');
  renderSidebarTickers();
  // Show sector dashboard if a specific sector is selected
  if (sec && FILE_LOADED) {
    const ws = document.getElementById('welcomeScreen');
    if (ws) ws.classList.add('hidden');
    renderSectorDashboard(sec);
  }
}

function filterSidebar() {
  sidebarQuery = document.getElementById('sidebarSearch').value.toLowerCase();
  renderSidebarTickers();
}

let screenerFilters = { price: '', score: '', value: '' };

function applyScreener() {
  screenerFilters.price = document.getElementById('filterPrice')?.value || '';
  screenerFilters.score = document.getElementById('filterScore')?.value || '';
  screenerFilters.value = document.getElementById('filterValue')?.value || '';
  renderSidebarTickers();
}

function clearScreener() {
  const fp = document.getElementById('filterPrice');
  const fs = document.getElementById('filterScore');
  const fv = document.getElementById('filterValue');
  if (fp) fp.value = '';
  if (fs) fs.value = '';
  if (fv) fv.value = '';
  screenerFilters = { price: '', score: '', value: '' };
  renderSidebarTickers();
}

function passesScreener(t) {
  const stk = STOCK[t] || {};
  const rows = ANNUAL[t];

  // Price filter
  if (screenerFilters.price) {
    const price = stk.price;
    if (price == null) return false;
    const p = screenerFilters.price;
    if (p === '0-50' && price > 50) return false;
    if (p === '50-100' && (price < 50 || price > 100)) return false;
    if (p === '100-250' && (price < 100 || price > 250)) return false;
    if (p === '250-500' && (price < 250 || price > 500)) return false;
    if (p === '500+' && price < 500) return false;
  }

  // Score filter
  if (screenerFilters.score && rows?.length) {
    const m = calcMetrics(rows);
    const s = calcScores(m).overall;
    const f = screenerFilters.score;
    if (f === '8+' && s < 8) return false;
    if (f === '6-8' && (s < 6 || s >= 8)) return false;
    if (f === '4-6' && (s < 4 || s >= 6)) return false;
    if (f === '0-4' && s >= 4) return false;
  }

  // Value filter
  if (screenerFilters.value && rows?.length) {
    const m = calcMetrics(rows);
    const f = screenerFilters.value;
    if (f === 'undervalued' && (stk.pe == null || stk.pe >= 15 || stk.pe <= 0)) return false;
    if (f === 'growth') {
      const revG = m.yoy.map(y => y.revenue_growth).filter(v => v != null);
      const avg = revG.length ? revG.reduce((a, b) => a + b, 0) / revG.length : 0;
      if (avg < 15) return false;
    }
    if (f === 'dividend' && (!stk.dividendYield || stk.dividendYield <= 0)) return false;
    if (f === 'lowdebt') {
      const dtes = m.lev.map(l => l.dte).filter(v => v != null);
      const avg = dtes.length ? dtes[dtes.length - 1] : null;
      if (avg == null || avg >= 0.5) return false;
    }
    if (f === 'profitable') {
      const margins = m.prof.map(p => p.net_margin).filter(v => v != null);
      const avg = margins.length ? margins[margins.length - 1] : null;
      if (avg == null || avg < 20) return false;
    }
  }

  return true;
}

function renderSidebarTickers() {
  const q = sidebarQuery;
  let list = TICKERS;
  if (activeSector) list = list.filter(t => STOCK[t]?.sector === activeSector);
  if (q) list = list.filter(t => t.toLowerCase().includes(q) || (STOCK[t]?.industry||'').toLowerCase().includes(q));

  // Apply screener filters
  const hasScreener = screenerFilters.price || screenerFilters.score || screenerFilters.value;
  if (hasScreener) list = list.filter(t => passesScreener(t));

  const countLabel = hasScreener ? ` <span style="font-size:10px;color:var(--accent)">(${list.length} matches)</span>` : '';
  const lblEl = document.getElementById('sidebarTickersLabel');
  if (lblEl) lblEl.innerHTML = (lang === 'ar' ? 'الشركات' : 'Companies') + countLabel;

  document.getElementById('sidebarTickers').innerHTML = list.slice(0,80).map(t =>
    `<div class="ticker-row${activeTicker===t?' active':''}" onclick="onTickerClick('${t}')">
      <span style="font-weight:700;font-size:13px">${t}</span>
      <span class="ticker-row-name">${STOCK[t]?.industry||''}${stk_price_tag(t)}</span>
    </div>`
  ).join('') + (list.length > 80 ? `<div style="padding:8px 10px;font-size:11px;color:var(--text3)">${list.length-80} more — search to filter</div>` : '');
}

// Show price in sidebar when screener active
function stk_price_tag(t) {
  const p = STOCK[t]?.price;
  if (!p || !screenerFilters.price) return '';
  return ` · $${p.toFixed(0)}`;
}

// ── HEADER SEARCH ─────────────────────────────────────────────────────────────
let dropIdx = -1;
let dropItems = [];

function onHdrSearch() {
  const q = document.getElementById('hdrSearch').value.toLowerCase().trim();
  const drop = document.getElementById('hdrDrop');
  if (!q || !TICKERS.length) { drop.classList.remove('open'); return; }

  dropItems = TICKERS.filter(t =>
    t.toLowerCase().includes(q) ||
    (STOCK[t]?.industry||'').toLowerCase().includes(q) ||
    (STOCK[t]?.sector||'').toLowerCase().includes(q)
  ).slice(0, 12);

  const isAr = lang === 'ar';
  // If the query looks like a plausible new ticker symbol (1-5 uppercase letters, alphanumeric),
  // offer to sync it from FMP — only when signed in.
  const upperQ = q.toUpperCase();
  const looksLikeTicker = /^[A-Z][A-Z.-]{0,5}$/.test(upperQ);
  const notInUniverse = looksLikeTicker && !TICKERS.includes(upperQ);
  const canSync = notInUniverse && isAiReady();

  let itemsHtml = dropItems.map((t, i) =>
    `<div class="hdr-search-item" onclick="pickTicker('${t}')" id="drop-item-${i}">
      <span class="hdr-search-ticker">${t}</span>
      <span class="hdr-search-meta">${STOCK[t]?.industry||''} · ${STOCK[t]?.sector||''}</span>
    </div>`
  ).join('');

  if (!itemsHtml && !canSync) {
    itemsHtml = '<div style="padding:12px 14px;font-size:13px;color:var(--text2)">No results</div>';
  }

  // Add "Sync new ticker" affordance when applicable
  if (canSync) {
    const hint = dropItems.length
      ? (isAr ? `لا ترى ${upperQ}؟` : `Don't see ${upperQ}?`)
      : (isAr ? `لم يتم العثور على ${upperQ}` : `${upperQ} not found in your universe`);
    itemsHtml += `<div class="hdr-search-sync" onclick="addTickerFromSearch('${upperQ}')" id="addTickerBtn">
      <div class="hdr-sync-icon">➕</div>
      <div class="hdr-sync-text">
        <div class="hdr-sync-title">${hint}</div>
        <div class="hdr-sync-sub">${isAr ? 'انقر لإضافته إلى الكون (~5 ثوانٍ)' : `Sync from FMP → add ${upperQ} to your universe (~5s)`}</div>
      </div>
    </div>`;
  } else if (notInUniverse && !isAiReady()) {
    itemsHtml += `<div class="hdr-search-sync hdr-sync-disabled">
      <div class="hdr-sync-icon">🔒</div>
      <div class="hdr-sync-text"><div class="hdr-sync-title">${isAr ? 'سجّل الدخول لإضافة رموز جديدة' : 'Sign in to add new tickers'}</div></div>
    </div>`;
  }

  drop.innerHTML = itemsHtml;
  drop.classList.add('open');
  dropIdx = -1;
}

// Triggered from the "Add to universe" button in the header search dropdown
async function addTickerFromSearch(ticker) {
  const isAr = lang === 'ar';
  const btn = document.getElementById('addTickerBtn');
  if (btn) {
    btn.style.pointerEvents = 'none';
    btn.querySelector('.hdr-sync-title').textContent = isAr ? `🔄 جارٍ المزامنة ${ticker}...` : `🔄 Syncing ${ticker}...`;
    btn.querySelector('.hdr-sync-sub').textContent = isAr ? 'جلب البيانات من FMP' : 'Fetching from FMP, please wait';
  }
  const result = await syncNewTickers(ticker);
  if (result?.synced?.length) {
    // Close dropdown and load the freshly-synced ticker's dashboard
    document.getElementById('hdrSearch').value = '';
    document.getElementById('hdrDrop').classList.remove('open');
    onTickerClick(ticker);
  } else {
    if (btn) {
      const errMsg = result?.errors?.[0]?.error || result?.error || 'unknown error';
      btn.querySelector('.hdr-sync-title').textContent = isAr ? `❌ فشلت المزامنة` : `❌ Sync failed`;
      btn.querySelector('.hdr-sync-sub').textContent = errMsg.slice(0, 80);
      btn.style.pointerEvents = '';
    }
  }
}

function onHdrKey(e) {
  const drop = document.getElementById('hdrDrop');
  if (!drop.classList.contains('open')) return;
  if (e.key === 'ArrowDown') { dropIdx = Math.min(dropIdx+1, dropItems.length-1); highlightDrop(); e.preventDefault(); }
  else if (e.key === 'ArrowUp') { dropIdx = Math.max(dropIdx-1, -1); highlightDrop(); e.preventDefault(); }
  else if (e.key === 'Enter' && dropIdx >= 0) { pickTicker(dropItems[dropIdx]); e.preventDefault(); }
  else if (e.key === 'Escape') { drop.classList.remove('open'); }
}

function highlightDrop() {
  document.querySelectorAll('.hdr-search-item').forEach((el, i) => {
    el.style.background = i === dropIdx ? 'var(--surface2)' : '';
  });
}

function openDrop() { if (document.getElementById('hdrSearch').value) onHdrSearch(); }

function pickTicker(t) {
  document.getElementById('hdrSearch').value = '';
  document.getElementById('hdrDrop').classList.remove('open');
  onTickerClick(t);
}

// Close drop on outside click
document.addEventListener('click', e => {
  if (!e.target.closest('.hdr-search-wrap')) document.getElementById('hdrDrop').classList.remove('open');
});

// ── METRICS ENGINE ────────────────────────────────────────────────────────────
function calcMetrics(rows) {
  rows = [...rows].sort((a,b) => a.year - b.year);
  const yoy=[], prof=[], lev=[], cf=[];
  for (let i=0; i<rows.length; i++) {
    const r=rows[i], p=rows[i-1];
    const pct=(a,b) => (a!=null&&b!=null&&b!==0) ? (a-b)/Math.abs(b)*100 : null;
    if (i>0) yoy.push({year:r.year,
      revenue_growth: pct(r.revenue,p.revenue),
      ni_growth:      pct(r.net_income,p.net_income),
      ocf_growth:     pct(r.operating_cash_flow,p.operating_cash_flow),
      eps_growth:     pct(r.eps_diluted,p.eps_diluted),
    });
    const {revenue:rv,net_income:ni,operating_income:oi,equity:eq,total_assets:ta,total_liabilities:tl,ebitda:eb} = r;
    prof.push({year:r.year,
      net_margin:  (rv&&ni&&rv!==0)  ? ni/rv*100  : null,
      op_margin:   (rv&&oi&&rv!==0)  ? oi/rv*100  : null,
      ebitda_margin:(rv&&eb&&rv!==0) ? eb/rv*100  : null,
      roe:         (eq&&ni&&eq!==0)  ? ni/eq*100  : null,
      roa:         (ta&&ni&&ta!==0)  ? ni/ta*100  : null,
    });
    const debt = r.total_debt || tl;
    lev.push({year:r.year,
      dte:  (eq&&debt&&eq!==0) ? debt/eq  : null,
      dta:  (ta&&debt&&ta!==0) ? debt/ta  : null,
    });
    cf.push({year:r.year,
      ocf:r.operating_cash_flow, fcf:r.free_cash_flow, capex:r.capital_expenditures,
      ocf_to_ni:(ni&&r.operating_cash_flow&&ni!==0) ? r.operating_cash_flow/ni : null,
    });
  }
  return {yoy, prof, lev, cf, rows, latest: rows[rows.length-1]||{}};
}

function calcScores(m) {
  const avg=(arr,k)=>{const v=arr.map(x=>x[k]).filter(v=>v!=null&&!isNaN(v));return v.length?v.reduce((a,b)=>a+b,0)/v.length:null};
  const avgRev=avg(m.yoy,'revenue_growth');
  let g=5;
  if(avgRev!=null){if(avgRev>=20)g=9.5;else if(avgRev>=15)g=8.5;else if(avgRev>=10)g=7.5;else if(avgRev>=5)g=6.5;else if(avgRev>=0)g=5;else if(avgRev>=-5)g=3.5;else g=2}
  const avgM=avg(m.prof,'net_margin'),avgR=avg(m.prof,'roe');
  let p=5;
  if(avgM!=null){if(avgM>=20)p=Math.min(p+2.5,10);else if(avgM>=10)p=Math.min(p+1.5,10);else if(avgM>=5)p=Math.min(p+.5,10);else if(avgM<0)p=Math.max(p-2,0)}
  if(avgR!=null){if(avgR>=20)p=Math.min(p+1.5,10);else if(avgR>=10)p=Math.min(p+.5,10);else if(avgR<0)p=Math.max(p-1,0)}
  const avgD=avg(m.lev,'dte');
  let h=5;
  if(avgD!=null){if(avgD<=.5)h=9;else if(avgD<=1)h=7.5;else if(avgD<=2)h=6;else if(avgD<=3)h=4.5;else if(avgD<=5)h=3;else h=1.5}
  const avgOCF=avg(m.cf,'ocf'),avgFCF=avg(m.cf,'fcf'),avgRat=avg(m.cf,'ocf_to_ni');
  let c=5;
  if(avgOCF!=null)c+=avgOCF>0?1.5:-2;
  if(avgFCF!=null&&avgFCF>0)c+=1.5;
  if(avgRat!=null){if(avgRat>=1.2)c+=1;else if(avgRat<.5)c-=1}
  c=Math.max(0,Math.min(10,c));
  const overall=Math.round(((g+p+h+c)/4)*10)/10;
  return{growth:Math.round(g*10)/10,profitability:Math.round(p*10)/10,health:Math.round(h*10)/10,cashflow:Math.round(c*10)/10,overall};
}

// ── INSIGHTS ──────────────────────────────────────────────────────────────────
function getSectorBenchmarks(ticker) {
  const sec = STOCK[ticker]?.sector;
  if (!sec) return null;
  const peers = TICKERS.filter(t => t !== ticker && STOCK[t]?.sector === sec && ANNUAL[t]?.length);
  if (peers.length < 3) return null;
  const peerMetrics = peers.map(t => {
    const m = calcMetrics(ANNUAL[t]);
    const s = calcScores(m);
    return { m, s };
  });
  const avg = (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
  return {
    sector: sec,
    count: peers.length + 1,
    avgOverall: avg(peerMetrics.map(p => p.s.overall).filter(v => v != null)),
    avgGrowth: avg(peerMetrics.flatMap(p => p.m.yoy.map(y => y.revenue_growth)).filter(v => v != null)),
    avgMargin: avg(peerMetrics.flatMap(p => p.m.prof.map(p2 => p2.net_margin)).filter(v => v != null)),
    avgDTE: avg(peerMetrics.flatMap(p => p.m.lev.map(l => l.dte)).filter(v => v != null)),
    avgROE: avg(peerMetrics.flatMap(p => p.m.prof.map(p2 => p2.roe)).filter(v => v != null)),
  };
}

function genInsights(ticker,m,scores) {
  const out={en:[],ar:[]};
  const add=(type,icon,en,ar)=>{out.en.push({type,icon,text:en});out.ar.push({type,icon,text:ar})};
  const bench = getSectorBenchmarks(ticker);

  // Revenue growth
  const revG=m.yoy.map(y=>y.revenue_growth).filter(v=>v!=null);
  if(revG.length){
    const avg=revG.reduce((a,b)=>a+b,0)/revG.length;
    if(avg>10) add('positive','📈',`${ticker} shows strong average revenue growth of ${avg.toFixed(1)}% per year.`,`يُظهر ${ticker} نمواً قوياً في الإيرادات بمتوسط ${avg.toFixed(1)}٪ سنوياً.`);
    else if(avg<0) add('warning','📉',`Revenue declining at ${Math.abs(avg).toFixed(1)}% annually — monitor market positioning.`,`تراجع الإيرادات بمعدل ${Math.abs(avg).toFixed(1)}٪ سنوياً — راقب التموضع السوقي.`);
    else add('neutral','➡️',`Moderate growth of ${avg.toFixed(1)}% — stable but may benefit from acceleration.`,`نمو معتدل ${avg.toFixed(1)}٪ — مستقر لكن يحتاج تسريعاً.`);
    // Sector context for growth
    if(bench?.avgGrowth!=null){
      const diff = avg - bench.avgGrowth;
      if(Math.abs(diff) > 3){
        const above = diff > 0;
        add(above?'positive':'warning', above?'🏆':'🔻',
          `Revenue growth is ${Math.abs(diff).toFixed(1)}pp ${above?'above':'below'} the ${bench.sector} sector average of ${bench.avgGrowth.toFixed(1)}%.`,
          `نمو الإيرادات ${above?'أعلى':'أقل'} بـ ${Math.abs(diff).toFixed(1)} نقطة من متوسط قطاع ${bench.sector} البالغ ${bench.avgGrowth.toFixed(1)}٪.`);
      }
    }
  }

  // Net margins
  const margins=m.prof.map(p=>p.net_margin).filter(v=>v!=null);
  if(margins.length){
    const avg=margins.reduce((a,b)=>a+b,0)/margins.length;
    if(avg>15) add('positive','💰',`Excellent net margins averaging ${avg.toFixed(1)}% — strong earnings retention.`,`هوامش صافية ممتازة بمتوسط ${avg.toFixed(1)}٪ — احتجاز قوي للأرباح.`);
    else if(avg<0) add('danger','🚨',`Net losses (${avg.toFixed(1)}% margin) — cash sustainability is critical.`,`خسائر صافية (هامش ${avg.toFixed(1)}٪) — استدامة النقد أمر بالغ الأهمية.`);
    else if(avg<5) add('warning','⚠️',`Thin net margins at ${avg.toFixed(1)}% — pricing or cost pressure likely.`,`هوامش رفيعة ${avg.toFixed(1)}٪ — ضغط تسعيري أو تكاليف مرتفعة.`);
    // Sector context for margins
    if(bench?.avgMargin!=null){
      const diff = avg - bench.avgMargin;
      if(Math.abs(diff) > 4){
        const above = diff > 0;
        add(above?'positive':'warning', above?'💎':'📉',
          `Net margin is ${Math.abs(diff).toFixed(1)}pp ${above?'above':'below'} the ${bench.sector} average of ${bench.avgMargin.toFixed(1)}%.`,
          `الهامش الصافي ${above?'أعلى':'أقل'} بـ ${Math.abs(diff).toFixed(1)} نقطة من متوسط ${bench.sector} البالغ ${bench.avgMargin.toFixed(1)}٪.`);
      }
    }
  }

  // Leverage
  const dtes=m.lev.map(l=>l.dte).filter(v=>v!=null);
  if(dtes.length){
    const avg=dtes.reduce((a,b)=>a+b,0)/dtes.length;
    if(avg>3) add('danger','⚠️',`High leverage — D/E of ${avg.toFixed(1)}x signals elevated financial risk.`,`رافعة مالية عالية — نسبة دين ${avg.toFixed(1)}x تشير إلى مخاطر مرتفعة.`);
    else if(avg<0.5) add('positive','🏦',`Conservative balance sheet — low D/E of ${avg.toFixed(1)}x provides flexibility.`,`ميزانية محافظة — نسبة دين منخفضة ${avg.toFixed(1)}x توفر مرونة.`);
    // Sector context for leverage
    if(bench?.avgDTE!=null){
      const diff = avg - bench.avgDTE;
      if(Math.abs(diff) > 0.5){
        const higher = diff > 0;
        add(higher?'warning':'positive', higher?'📊':'🛡️',
          `D/E ratio of ${avg.toFixed(2)}x is ${Math.abs(diff).toFixed(2)}x ${higher?'higher':'lower'} than the ${bench.sector} average of ${bench.avgDTE.toFixed(2)}x.`,
          `نسبة الدين ${avg.toFixed(2)}x ${higher?'أعلى':'أقل'} بـ ${Math.abs(diff).toFixed(2)}x من متوسط ${bench.sector} البالغ ${bench.avgDTE.toFixed(2)}x.`);
      }
    }
  }

  // OCF
  const ocfs=m.cf.map(c=>c.ocf).filter(v=>v!=null);
  if(ocfs.length){
    const avg=ocfs.reduce((a,b)=>a+b,0)/ocfs.length;
    if(avg>0) add('positive','💵',`Consistently positive operating cash flow — real cash from core operations.`,`تدفق نقدي تشغيلي إيجابي باستمرار — نقد حقيقي من العمليات الأساسية.`);
    else add('warning','🔴',`Negative operating cash flow — long-term self-funding at risk.`,`تدفق نقدي تشغيلي سلبي — التمويل الذاتي على المدى البعيد في خطر.`);
  }

  // FCF
  const fcfs=m.cf.map(c=>c.fcf).filter(v=>v!=null);
  if(fcfs.length&&fcfs.reduce((a,b)=>a+b,0)/fcfs.length>0)
    add('positive','🟢',`Positive free cash flow — able to fund growth or return capital.`,`تدفق نقدي حر إيجابي — قادر على تمويل النمو أو إعادة رأس المال.`);

  // ROE sector context
  if(bench?.avgROE!=null){
    const roes = m.prof.map(p=>p.roe).filter(v=>v!=null);
    if(roes.length){
      const avgROE = roes.reduce((a,b)=>a+b,0)/roes.length;
      const diff = avgROE - bench.avgROE;
      if(Math.abs(diff) > 5){
        const above = diff > 0;
        add(above?'positive':'warning', above?'⚡':'📉',
          `ROE of ${avgROE.toFixed(1)}% is ${Math.abs(diff).toFixed(1)}pp ${above?'above':'below'} the ${bench.sector} average of ${bench.avgROE.toFixed(1)}%.`,
          `العائد على الملكية ${avgROE.toFixed(1)}٪ ${above?'أعلى':'أقل'} بـ ${Math.abs(diff).toFixed(1)} نقطة من متوسط ${bench.sector} البالغ ${bench.avgROE.toFixed(1)}٪.`);
      }
    }
  }

  // Overall score sector context
  if(bench?.avgOverall!=null){
    const diff = scores.overall - bench.avgOverall;
    if(Math.abs(diff) > 1){
      const above = diff > 0;
      add(above?'positive':'warning', above?'🏅':'📋',
        `Overall score ${scores.overall}/10 is ${Math.abs(diff).toFixed(1)} points ${above?'above':'below'} the ${bench.sector} sector average of ${bench.avgOverall.toFixed(1)}/10.`,
        `الدرجة الإجمالية ${scores.overall}/10 ${above?'أعلى':'أقل'} بـ ${Math.abs(diff).toFixed(1)} نقطة من متوسط قطاع ${bench.sector} البالغ ${bench.avgOverall.toFixed(1)}/10.`);
    }
  }

  return out;
}

// ── DeepSeek V4 API INTEGRATION (via Supabase proxy) ────────────────────────
// Available models — toggle cycles in this order. First is default.
// Change ids to any OpenAI model your project has access to (gpt-4o, gpt-4.1-mini, gpt-5-mini, etc.)
const AI_MODELS = [
  { id: 'gpt-4o-mini',       label: '🤖 GPT-4o mini', short: 'GPT' },
  { id: 'deepseek-v4-flash', label: '🤖 DeepSeek V4', short: 'DSV4' },
];
let DEEPSEEK_MODEL = AI_MODELS[0].id;
let aiAbortController = null;

function getProxyUrl() { return SUPABASE_URL + '/functions/v1/rapid-worker'; }
function getSeedKgUrl() { return SUPABASE_URL + '/functions/v1/seed-kg'; }
function getSyncTickerUrl() { return SUPABASE_URL + '/functions/v1/sync-ticker'; }
function getProxyHeaders() {
  return {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
    'x-auth-hash': USER_AUTH_HASH || '',
  };
}
function isAiReady() { return !!USER_AUTH_HASH; }

function getModelMeta(id) {
  return AI_MODELS.find(m => m.id === id) || AI_MODELS[0];
}
function toggleModel() {
  const idx = AI_MODELS.findIndex(m => m.id === DEEPSEEK_MODEL);
  DEEPSEEK_MODEL = AI_MODELS[(idx + 1) % AI_MODELS.length].id;
  const meta = getModelMeta(DEEPSEEK_MODEL);
  const btn = document.getElementById('modelBtn');
  if (btn) btn.textContent = meta.label;
  const badge = document.querySelector('.ai-model-badge');
  if (badge) badge.textContent = meta.short;
  scheduleMemorySave();
}

// ── AGENTIC SYSTEM ───────────────────────────────────────────────────────────
const AI_MAX_RETRIES = 2;
const AI_MAX_TOKENS = 2048;
const AI_MAX_CONTINUATIONS = 2;
let agentMemory = {}; // ticker -> {analysis, timestamp, findings}

// Helper: build a data snapshot string for a given ticker
function buildTickerSnapshot(ticker) {
  const rows = ANNUAL[ticker];
  if (!rows || !rows.length) return null;
  const m = calcMetrics(rows);
  const scores = calcScores(m);
  const stk = STOCK[ticker] || {};
  const lat = m.latest || {};
  const yoyData = m.yoy.map(y => `${y.year}: Rev ${y.revenue_growth!=null?y.revenue_growth.toFixed(1)+'%':'N/A'}, NI ${y.ni_growth!=null?y.ni_growth.toFixed(1)+'%':'N/A'}`).join('\n');
  const profData = m.prof.map(p => `${p.year}: NetM ${p.net_margin!=null?p.net_margin.toFixed(1)+'%':'N/A'}, OpM ${p.op_margin!=null?p.op_margin.toFixed(1)+'%':'N/A'}, ROE ${p.roe!=null?p.roe.toFixed(1)+'%':'N/A'}`).join('\n');
  const levData = m.lev.map(l => `${l.year}: D/E ${l.dte!=null?l.dte.toFixed(2)+'x':'N/A'}`).join('\n');
  const cfData = m.cf.map(c => `${c.year}: OCF ${c.ocf!=null?fM(c.ocf):'N/A'}, FCF ${c.fcf!=null?fM(c.fcf):'N/A'}`).join('\n');
  const est = ESTIMATES[ticker];
  let estStr = '';
  if (est && est.length) {
    const nextEst = est.filter(e => new Date(e.date) > new Date()).sort((a,b) => a.date.localeCompare(b.date))[0] || est[est.length - 1];
    if (nextEst) {
      estStr = `\nAnalyst Estimates (${nextEst.date?.slice(0,4)||''}): Rev ${nextEst.revenueAvg?fM(nextEst.revenueAvg):'N/A'}, EPS $${nextEst.epsAvg?.toFixed(2)||'N/A'} (range $${nextEst.epsLow?.toFixed(2)||'?'}-$${nextEst.epsHigh?.toFixed(2)||'?'}), NI ${nextEst.netIncomeAvg?fM(nextEst.netIncomeAvg):'N/A'}, ${nextEst.numAnalystsEps||'?'} analysts`;
    }
  }
  return {m, scores, stk, summary: `${ticker} | Sector: ${stk.sector||'N/A'} | Industry: ${stk.industry||'N/A'}
Scores: Growth ${scores.growth}, Profitability ${scores.profitability}, Health ${scores.health}, CashFlow ${scores.cashflow}, Overall ${scores.overall}/10
${stk.price?'Price: $'+stk.price.toFixed(2)+' ':''}${stk.change!=null?'Day Change: '+(stk.change>=0?'+':'')+stk.change.toFixed(2)+' ('+(stk.changePct>=0?'+':'')+stk.changePct?.toFixed(2)+'%) ':''}${stk.pe?'P/E: '+stk.pe.toFixed(1)+' ':''}${stk.marketCap?'MCap: '+fM(stk.marketCap)+' ':''}${stk.yearHigh?'52w High: $'+stk.yearHigh.toFixed(2)+' ':''}${stk.yearLow?'52w Low: $'+stk.yearLow.toFixed(2)+' ':''}
YoY: ${yoyData}
Profitability: ${profData}
Leverage: ${levData}
CashFlow: ${cfData}
Latest: Rev ${fM(lat.revenue)}, NI ${fM(lat.net_income)}, Assets ${fM(lat.total_assets)}, Equity ${fM(lat.equity)}, EPS ${lat.eps_diluted!=null?'$'+lat.eps_diluted.toFixed(2):'N/A'}${estStr}`};
}

// Find peers: same sector tickers in the dataset
function findPeers(ticker, maxPeers) {
  const sector = STOCK[ticker]?.sector;
  if (!sector) return [];
  return TICKERS.filter(t => t !== ticker && STOCK[t]?.sector === sector && ANNUAL[t]?.length).slice(0, maxPeers || 5);
}

// UI helpers for agent panel
function agentConv() { return document.getElementById('agentConv'); }

function addAgentStep(type, icon, label, body) {
  const conv = agentConv(); if (!conv) return;
  const div = document.createElement('div');
  div.className = `agent-step ${type}`;
  div.innerHTML = `<div class="step-hdr"><span class="step-icon">${icon}</span><span class="step-label ${type}">${label}</span></div><div class="step-body">${body}</div>`;
  conv.appendChild(div);
  conv.scrollTop = conv.scrollHeight;
  return div;
}

function addAgentStreamStep(type, icon, label) {
  const conv = agentConv(); if (!conv) return;
  const div = document.createElement('div');
  div.className = `agent-step ${type}`;
  div.innerHTML = `<div class="step-hdr"><span class="step-icon">${icon}</span><span class="step-label ${type}">${label}</span></div><div class="step-body"></div>`;
  conv.appendChild(div);
  conv.scrollTop = conv.scrollHeight;
  return div.querySelector('.step-body');
}

function setAgentStatus(active, text) {
  const dot = document.querySelector('.agent-status-dot');
  const txt = document.getElementById('agentStatusTxt');
  if (dot) { dot.className = `agent-status-dot ${active ? 'active' : 'idle'}`; }
  if (txt) txt.textContent = text;
}

function showAgentActions(actions) {
  const el = document.getElementById('agentActions');
  if (!el) return;
  el.classList.remove('hidden');
  const lbl = lang === 'ar' ? 'الخطوات المقترحة:' : 'Suggested next:';
  el.innerHTML = `<span class="agent-actions-label">${lbl}</span>` +
    actions.map(a => `<button class="action-chip" onclick="${a.onclick}">${a.label}</button>`).join('');
}

function clearAgentPanel() {
  const conv = agentConv(); if (conv) conv.innerHTML = '';
  const acts = document.getElementById('agentActions'); if (acts) { acts.innerHTML = ''; acts.classList.add('hidden'); }
  setAgentStatus(false, lang === 'ar' ? 'الوكيل في وضع الاستعداد' : 'Agent standby');
}

// API call with streaming into a DOM element
async function streamToElement(messages, el, signal) {
  let fullText = '';
  let finishReason = null;

  const response = await fetch(getProxyUrl(), {
    method: 'POST',
    headers: getProxyHeaders(),
    body: JSON.stringify({ model: DEEPSEEK_MODEL, messages, stream: true, max_tokens: AI_MAX_TOKENS, temperature: 0.7 }),
    signal
  });

  if (!response.ok) {
    const errData = await response.json().catch(() => ({}));
    throw new Error(errData.error?.message || `HTTP ${response.status}`);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() || '';
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed === 'data: [DONE]' || !trimmed.startsWith('data: ')) continue;
      try {
        const json = JSON.parse(trimmed.slice(6));
        const c = json.choices?.[0];
        if (c?.delta?.content) { fullText += c.delta.content; el.innerHTML = fmtAI(fullText); const sc = el.closest('.chat-messages') || agentConv(); if (sc) sc.scrollTop = sc.scrollHeight; }
        if (c?.finish_reason) finishReason = c.finish_reason;
      } catch(e) {}
    }
  }
  // Process remaining buffer
  if (buffer.trim() && buffer.trim() !== 'data: [DONE]' && buffer.trim().startsWith('data: ')) {
    try {
      const json = JSON.parse(buffer.trim().slice(6));
      if (json.choices?.[0]?.delta?.content) { fullText += json.choices[0].delta.content; el.innerHTML = fmtAI(fullText); }
      if (json.choices?.[0]?.finish_reason) finishReason = json.choices[0].finish_reason;
    } catch(e) {}
  }

  // Handle truncation — auto-continue
  let cont = 0;
  while (finishReason === 'length' && cont < AI_MAX_CONTINUATIONS) {
    cont++;
    const contMsg = [...messages, { role: 'assistant', content: fullText },
      { role: 'user', content: lang === 'ar' ? 'أكمل من حيث توقفت، لا تكرر.' : 'Continue from where you stopped, no repetition.' }];
    const r2 = await fetch(getProxyUrl(), {
      method: 'POST', headers: getProxyHeaders(),
      body: JSON.stringify({ model: DEEPSEEK_MODEL, messages: contMsg, stream: true, max_tokens: AI_MAX_TOKENS, temperature: 0.7 }),
      signal
    });
    if (!r2.ok) break;
    const rd2 = r2.body.getReader(); let buf2 = ''; finishReason = null;
    while (true) {
      const { done, value } = await rd2.read(); if (done) break;
      buf2 += decoder.decode(value, { stream: true });
      const ls = buf2.split('\n'); buf2 = ls.pop() || '';
      for (const l of ls) {
        const t = l.trim(); if (!t || t === 'data: [DONE]' || !t.startsWith('data: ')) continue;
        try { const j = JSON.parse(t.slice(6)); if (j.choices?.[0]?.delta?.content) { fullText += j.choices[0].delta.content; el.innerHTML = fmtAI(fullText); agentConv()?.scrollTo(0, agentConv().scrollHeight); } if (j.choices?.[0]?.finish_reason) finishReason = j.choices[0].finish_reason; } catch(e) {}
      }
    }
  }
  return fullText;
}

// Retry wrapper
async function streamWithRetry(messages, el, signal) {
  for (let attempt = 0; attempt <= AI_MAX_RETRIES; attempt++) {
    if (signal.aborted) return '';
    if (attempt > 0) {
      el.textContent = lang === 'ar' ? `إعادة المحاولة (${attempt})...` : `Retrying (${attempt})...`;
      await new Promise(r => setTimeout(r, 1500 * attempt));
    }
    try {
      const text = await streamToElement(messages, el, signal);
      if (text.trim()) return text;
    } catch(e) {
      if (e.name === 'AbortError') return '';
      if (attempt === AI_MAX_RETRIES) throw e;
    }
  }
  return '';
}

// ── AGENTIC LOOP ─────────────────────────────────────────────────────────────
async function agentStart(ticker) {
  if (!isAiReady()) { showPinModal(); return; }
  if (!currentDash) return;

  // Cancel previous
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  // Toggle buttons
  const genBtn = document.getElementById('aiGenBtn');
  const regenBtn = document.getElementById('aiRegenBtn');
  if (genBtn) genBtn.classList.add('hidden');
  if (regenBtn) { regenBtn.classList.remove('hidden'); regenBtn.disabled = true; }

  clearAgentPanel();
  setAgentStatus(true, lang === 'ar' ? 'يحلل البيانات...' : 'Analysing data...');

  const snap = buildTickerSnapshot(ticker);
  if (!snap) { addAgentStep('warning', '⚠️', 'Error', 'No data for this ticker.'); setAgentStatus(false, 'Error'); return; }

  try {
    // ── STEP 1: Core Analysis ──
    const aiLength = document.getElementById('aiLengthSelect')?.value || 'standard';
    const lengthGuide = {
      brief: { en: 'Write 2 short paragraphs (80-120 words). Be extremely concise — key takeaway and one risk only.', ar: 'اكتب فقرتين قصيرتين (80-120 كلمة). كن مختصراً جداً — النقطة الرئيسية وخطر واحد فقط.' },
      standard: { en: 'Write 3-4 flowing paragraphs (200-300 words): overall assessment, key strengths, risks, and outlook.', ar: 'اكتب 3-4 فقرات (200-300 كلمة): التقييم العام، نقاط القوة، المخاطر، والنظرة المستقبلية.' },
      deep: { en: 'Write a comprehensive 5-6 paragraph analysis (400-500 words): detailed assessment, strengths with evidence, all risks and mitigants, competitive position, valuation context, and forward outlook with specific catalysts.', ar: 'اكتب تحليلاً شاملاً من 5-6 فقرات (400-500 كلمة): تقييم مفصل، نقاط القوة مع أدلة، جميع المخاطر وعوامل التخفيف، الموقف التنافسي، سياق التقييم، والنظرة المستقبلية مع محفزات محددة.' },
    };
    const sysPrompt = lang === 'ar'
      ? `أنت وكيل تحليل مالي ذكي. ${lengthGuide[aiLength].ar} كن دقيقاً بالأرقام. لا تستخدم عناوين أو نقاط أو ماركداون. أنهِ بجملة ختامية.`
      : `You are an AI financial analysis agent. ${lengthGuide[aiLength].en} Be specific with numbers. No bullet points, headers, or markdown. End with a clear concluding sentence.`;
    
    addAgentStep('thinking', '🧠', lang==='ar'?'يفكر...':'Thinking...', lang==='ar'?'يراجع البيانات المالية ويحسب المؤشرات الرئيسية...':'Reviewing financial data and computing key indicators...');
    if (signal.aborted) return;
    await new Promise(r => setTimeout(r, 400));

    setAgentStatus(true, lang === 'ar' ? 'يولّد التحليل الأساسي...' : 'Generating core analysis...');
    const analysisEl = addAgentStreamStep('analysis', '📊', lang==='ar'?'التحليل الأساسي':'Core Analysis');
    const analysisText = await streamWithRetry([
      { role: 'system', content: sysPrompt },
      { role: 'user', content: snap.summary }
    ], analysisEl, signal);

    if (!analysisText.trim()) throw new Error(lang==='ar'?'لم يُرسل النموذج رداً':'Model returned empty');

    // Store in memory
    agentMemory[ticker] = { analysis: analysisText, timestamp: Date.now(), scores: snap.scores };
    scheduleMemorySave(); // Auto-save to cloud

    // ── STEP 1.5: Fetch and display news context ──
    let newsContext = '';
    if (FMP_QUOTE_KEY) {
      setAgentStatus(true, lang === 'ar' ? 'يبحث عن الأخبار...' : 'Scanning recent news...');
      let news = await fetchLiveNews(ticker).catch(() => []);
      if (news && news.length) {
        // Score sentiment
        if (isAiReady()) {
          setAgentStatus(true, lang === 'ar' ? 'يحلل مشاعر الأخبار...' : 'Analysing news sentiment...');
          news = await analyseNewsSentiment(news, ticker);
          // Also update the news feed UI
          renderNewsSection(news);
        }

        const newsHeadlines = news.slice(0, 5).map(n => `• [${(n.sentiment||'neutral').toUpperCase()}] ${n.title} (${n.publisher || n.site}, ${n.publishedDate?.slice(0,10)||''})`).join('\n');
        const bull = news.filter(n => n.sentiment === 'bullish').length;
        const bear = news.filter(n => n.sentiment === 'bearish').length;
        const sentimentSummary = bull > bear ? 'Overall sentiment: BULLISH' : bear > bull ? 'Overall sentiment: BEARISH' : 'Overall sentiment: MIXED/NEUTRAL';
        newsContext = `\n\nRecent headlines (sentiment-scored):\n${newsHeadlines}\n${sentimentSummary} (${bull} bullish, ${bear} bearish, ${news.length - bull - bear} neutral)`;

        // Store in memory
        agentMemory[ticker].news = newsHeadlines;
        agentMemory[ticker].sentiment = bull - bear;

        const isAr = lang === 'ar';
        const sentColor = bull > bear ? 'var(--green)' : bear > bull ? 'var(--red)' : 'var(--text3)';
        const sentLabel = bull > bear ? (isAr ? 'إيجابي' : 'Bullish') : bear > bull ? (isAr ? 'سلبي' : 'Bearish') : (isAr ? 'محايد' : 'Neutral');
        addAgentStep('observation', '📰', isAr ? 'أخبار ومشاعر السوق' : 'News & Sentiment',
          `<div style="margin-bottom:8px;display:flex;align-items:center;gap:8px"><span style="font-size:12px;font-weight:700;color:${sentColor}">${sentLabel}</span><span style="font-size:10px;color:var(--text3)">${bull}🟢 ${news.length-bull-bear}⚪ ${bear}🔴</span></div>` +
          news.slice(0, 3).map(n => `<div style="margin-bottom:6px;display:flex;align-items:flex-start;gap:8px"><span class="news-sentiment ${n.sentiment}" style="flex-shrink:0;margin-top:2px">${n.sentiment === 'bullish' ? '▲' : n.sentiment === 'bearish' ? '▼' : '—'}</span><div><a href="${n.url}" target="_blank" style="color:var(--accent);text-decoration:none;font-size:12px;font-weight:500">${n.title}</a><div style="font-size:10px;color:var(--text3)">${n.publisher || ''} · ${n.publishedDate?.slice(0,10)||''}</div></div></div>`).join(''));
      }
    }

    // ── STEP 2: Auto-detect concerns & observations ──
    if (signal.aborted) return;
    setAgentStatus(true, lang === 'ar' ? 'يفحص المخاطر...' : 'Scanning for risks...');
    await new Promise(r => setTimeout(r, 300));

    const concerns = [];
    const avgDte = snap.m.lev.map(l=>l.dte).filter(v=>v!=null);
    const latestDte = avgDte.length ? avgDte[avgDte.length-1] : null;
    if (latestDte != null && latestDte > 2) concerns.push({icon:'⚠️', type:'warning', text: lang==='ar' ? `رافعة مالية مرتفعة: نسبة الدين/الملكية ${latestDte.toFixed(2)}x — أعلى من عتبة الأمان 2x` : `High leverage detected: D/E ratio at ${latestDte.toFixed(2)}x — above the 2x safety threshold`});

    const margins = snap.m.prof.map(p=>p.net_margin).filter(v=>v!=null);
    if (margins.length >= 2) {
      const recent = margins.slice(-2);
      if (recent[1] < recent[0] && recent[1] < 5) concerns.push({icon:'📉', type:'warning', text: lang==='ar' ? `هوامش الربح تتراجع: من ${recent[0].toFixed(1)}٪ إلى ${recent[1].toFixed(1)}٪ — ضغط تنافسي محتمل` : `Margins declining: from ${recent[0].toFixed(1)}% to ${recent[1].toFixed(1)}% — possible competitive pressure`});
    }

    const revGrowths = snap.m.yoy.map(y=>y.revenue_growth).filter(v=>v!=null);
    if (revGrowths.length >= 2) {
      const lastTwo = revGrowths.slice(-2);
      if (lastTwo[1] < lastTwo[0] * 0.5 && lastTwo[0] > 5) concerns.push({icon:'🔻', type:'warning', text: lang==='ar' ? `تباطؤ حاد في النمو: من ${lastTwo[0].toFixed(1)}٪ إلى ${lastTwo[1].toFixed(1)}٪` : `Sharp growth deceleration: from ${lastTwo[0].toFixed(1)}% to ${lastTwo[1].toFixed(1)}%`});
    }

    const fcfs = snap.m.cf.map(c=>c.fcf).filter(v=>v!=null);
    const latestFCF = fcfs.length ? fcfs[fcfs.length-1] : null;
    if (latestFCF != null && latestFCF < 0) concerns.push({icon:'💸', type:'warning', text: lang==='ar' ? `تدفق نقدي حر سلبي: ${fM(latestFCF)} — قد يحتاج تمويل خارجي` : `Negative FCF: ${fM(latestFCF)} — may need external financing`});

    if (snap.scores.overall >= 8) concerns.push({icon:'🏆', type:'observation', text: lang==='ar' ? `أداء استثنائي: الدرجة الإجمالية ${snap.scores.overall}/10 تضعها في الشريحة العليا` : `Exceptional performer: Overall score ${snap.scores.overall}/10 places it in the top tier`});

    if (concerns.length) {
      for (const c of concerns) { addAgentStep(c.type === 'observation' ? 'observation' : 'warning', c.icon, c.type === 'observation' ? (lang==='ar'?'ملاحظة':'Observation') : (lang==='ar'?'تنبيه':'Alert'), c.text); }
    } else {
      addAgentStep('observation', '✅', lang==='ar'?'ملاحظة':'Observation', lang==='ar'?'لم يتم رصد مخاطر كبيرة. المؤشرات المالية ضمن النطاق الطبيعي.':'No major red flags detected. Financial metrics are within normal ranges.');
    }

    // ── STEP 3: Peer discovery ──
    if (signal.aborted) return;
    const peers = findPeers(ticker, 4);
    if (peers.length >= 2) {
      setAgentStatus(true, lang === 'ar' ? 'يقارن مع الأقران...' : 'Comparing with peers...');
      await new Promise(r => setTimeout(r, 300));

      const peerData = peers.map(t => {
        const ps = buildTickerSnapshot(t);
        return ps ? { ticker: t, ...ps.scores, sector: STOCK[t]?.sector, industry: STOCK[t]?.industry } : null;
      }).filter(Boolean);

      if (peerData.length) {
        let tblHtml = `<table class="agent-peer-tbl"><thead><tr><th>${lang==='ar'?'الشركة':'Company'}</th><th>${lang==='ar'?'النمو':'Growth'}</th><th>${lang==='ar'?'الربحية':'Profit'}</th><th>${lang==='ar'?'الصحة':'Health'}</th><th>${lang==='ar'?'التدفق':'CF'}</th><th>${lang==='ar'?'الإجمالي':'Overall'}</th></tr></thead><tbody>`;
        // Current ticker first
        tblHtml += `<tr class="current"><td><strong>${ticker}</strong></td><td>${snap.scores.growth}</td><td>${snap.scores.profitability}</td><td>${snap.scores.health}</td><td>${snap.scores.cashflow}</td><td><strong>${snap.scores.overall}</strong></td></tr>`;
        for (const p of peerData) {
          const oC = p.overall > snap.scores.overall ? 'pos' : p.overall < snap.scores.overall ? 'neg' : '';
          tblHtml += `<tr><td>${p.ticker}</td><td>${p.growth}</td><td>${p.profitability}</td><td>${p.health}</td><td>${p.cashflow}</td><td class="${oC}"><strong>${p.overall}</strong></td></tr>`;
        }
        tblHtml += '</tbody></table>';

        addAgentStep('comparison', '🔄', lang==='ar'?`مقارنة الأقران (${STOCK[ticker]?.sector})`:`Peer Comparison (${STOCK[ticker]?.sector})`, tblHtml);

        // ── STEP 4: AI peer interpretation ──
        if (signal.aborted) return;
        setAgentStatus(true, lang === 'ar' ? 'يحلل المقارنة...' : 'Interpreting comparison...');
        const peerSummaries = peerData.map(p => `${p.ticker}: Overall ${p.overall}/10, Growth ${p.growth}, Prof ${p.profitability}, Health ${p.health}`).join('\n');
        const peerSysPrompt = lang === 'ar'
          ? 'أنت محلل مالي. قارن هذه الشركة مع أقرانها في فقرة واحدة مختصرة (60-80 كلمة). أذكر أين تتفوق وأين تتأخر. لا عناوين ولا نقاط.'
          : 'You are a financial analyst. Compare this company vs its peers in one concise paragraph (60-80 words). Highlight where it leads and where it lags. No headers or bullets.';
        const compEl = addAgentStreamStep('comparison', '🤖', lang==='ar'?'تحليل المقارنة':'Comparative Insight');
        await streamWithRetry([
          { role: 'system', content: peerSysPrompt },
          { role: 'user', content: `${ticker} scores: Overall ${snap.scores.overall}, Growth ${snap.scores.growth}, Profitability ${snap.scores.profitability}, Health ${snap.scores.health}, CashFlow ${snap.scores.cashflow}.\nPeers:\n${peerSummaries}` }
        ], compEl, signal);
      }
    }

    // ── STEP 5: Final verdict with memory context ──
    if (signal.aborted) return;
    setAgentStatus(true, lang === 'ar' ? 'يصيغ الخلاصة...' : 'Forming verdict...');
    await new Promise(r => setTimeout(r, 300));

    // Check if we've analysed other tickers this session for cross-referencing
    const priorTickers = Object.keys(agentMemory).filter(t => t !== ticker);
    let memoryContext = '';
    if (priorTickers.length) {
      memoryContext = lang === 'ar'
        ? `\n\nملاحظة: سبق تحليل هذه الشركات في هذه الجلسة: ${priorTickers.map(t => `${t} (درجة ${agentMemory[t].scores.overall}/10)`).join('، ')}. إذا كانت المقارنة مفيدة، اذكرها باختصار.`
        : `\n\nNote: These companies were analysed earlier this session: ${priorTickers.map(t => `${t} (score ${agentMemory[t].scores.overall}/10)`).join(', ')}. If a comparison is useful, mention it briefly.`;
    }

    const verdictSys = lang === 'ar'
      ? 'أنت وكيل مالي. اكتب خلاصة نهائية في 2-3 جمل. ابدأ بـ "الخلاصة:" متبوعاً بتوصيتك. استند إلى التحليل المالي والمخاطر المكتشفة والأخبار الأخيرة وتقديرات المحللين. كن حاسماً ومحدداً.'
      : 'You are a financial agent. Write a final verdict in 2-3 sentences. Start with "Verdict:" followed by your recommendation. Base it on ALL inputs: the financial analysis, detected risks, recent news headlines, analyst estimates, and peer comparison. Reference specific data points. Be decisive.';
    const verdictEl = addAgentStreamStep('conclusion', '⚡', lang==='ar'?'خلاصة الوكيل':'Agent Verdict');
    await streamWithRetry([
      { role: 'system', content: verdictSys },
      { role: 'user', content: `Based on ${ticker} analysis: ${analysisText.slice(0, 400)}... Scores: Overall ${snap.scores.overall}/10. ${concerns.length ? 'Concerns: ' + concerns.map(c=>c.text).join('; ') : 'No major concerns.'}${newsContext}${memoryContext}` }
    ], verdictEl, signal);

    // ── STEP 6: SKEPTIC LAYER (Risk Strategist) ──
    // A second AI pass that challenges the verdict with 6 tests
    if (signal.aborted) return;
    setAgentStatus(true, lang === 'ar' ? '🧐 المشكك يراجع التحليل...' : '🧐 Skeptic reviewing analysis...');
    await new Promise(r => setTimeout(r, 400));

    const verdictText = verdictEl.textContent || '';
    const skepticSys = lang === 'ar'
      ? `أنت محلل مخاطر متشكك. مهمتك تحدي كل استنتاج. لديك 6 اختبارات:
1. تعدد المصادر: هل الاستنتاج مبني على مصادر متعددة أم تغريدة واحدة؟
2. التناسب: هل حجم التأثير المتوقع منطقي مقارنة بحجم السبب؟
3. حداثة البيانات: هل البيانات المستخدمة حديثة أم قديمة؟
4. المسار السببي: هل هناك علاقة سببية حقيقية أم مجرد ارتباط؟
5. التحيز التأكيدي: هل التحليل تجاهل أدلة معاكسة؟
6. المخاطر المخفية: ما الذي يمكن أن يسوء ولم يُذكر؟

أعطِ درجة ثقة من 1-10 وصنّف: "موثوق" أو "يحتاج حذر" أو "مرفوض".
اكتب 3-4 جمل فقط. ابدأ بـ "الثقة: X/10 —" ثم تقييمك.`
      : `You are a skeptical risk strategist. Your job is to challenge every conclusion. Run 6 tests:
1. SOURCE COUNT: Is the conclusion backed by multiple data sources, or based on a single data point?
2. MAGNITUDE CHECK: Does the predicted impact make sense relative to the cause? (e.g., a 0.2% move causing sector collapse = nonsense)
3. DATA FRESHNESS: Is the analysis using current data, or are the financials stale?
4. CAUSAL PATH: Is there a genuine cause-and-effect link, or just correlation?
5. CONFIRMATION BIAS: Did the analysis ignore contradicting evidence?
6. HIDDEN RISKS: What could go wrong that wasn't mentioned?

Assign a confidence score 1-10 and classify as: "VALIDATED", "CAUTION", or "REJECTED".
Write 3-4 sentences only. Start with "Confidence: X/10 —" then your assessment. Be specific about which tests passed or failed.`;

    const skepticEl = addAgentStreamStep('warning', '🧐', lang==='ar'?'تقييم المشكك':'Skeptic Assessment');
    const skepticText = await streamWithRetry([
      { role: 'system', content: skepticSys },
      { role: 'user', content: `VERDICT TO CHALLENGE:\n${verdictText}\n\nUNDERLYING DATA:\nScores: Overall ${snap.scores.overall}/10, Growth ${snap.scores.growth}, Profitability ${snap.scores.profitability}, Health ${snap.scores.health}\n${concerns.length ? 'Detected risks: ' + concerns.map(c=>c.text).join('; ') : 'No auto-detected risks.'}\n${newsContext ? 'News context: ' + newsContext : 'No recent news available.'}\nAnalyst estimates available: ${ESTIMATES[ticker]?.length ? 'Yes' : 'No'}\nPeer comparison done: ${peers.length >= 2 ? 'Yes, ' + peers.length + ' peers' : 'No peers found'}\nData source: Supabase cloud (quarterly sync)` }
    ], skepticEl, signal);

    // Extract confidence score and add visual indicator
    const confMatch = skepticText.match(/(\d+)\s*\/\s*10/);
    const confScore = confMatch ? parseInt(confMatch[1]) : null;
    if (confScore !== null && skepticEl) {
      const confColor = confScore >= 7 ? 'var(--green)' : confScore >= 4 ? 'var(--yellow)' : 'var(--red)';
      const confLabel = confScore >= 7 ? (lang==='ar'?'موثوق':'VALIDATED') : confScore >= 4 ? (lang==='ar'?'يحتاج حذر':'CAUTION') : (lang==='ar'?'مرفوض':'REJECTED');
      skepticEl.insertAdjacentHTML('afterend', `<div style="display:flex;align-items:center;gap:10px;margin-top:8px;padding:8px 12px;border-radius:8px;background:${confColor}15;border:1px solid ${confColor}30">
        <div style="font-size:24px;font-weight:700;color:${confColor};font-family:var(--display)">${confScore}/10</div>
        <div>
          <div style="font-size:11px;font-weight:700;color:${confColor};letter-spacing:.08em">${confLabel}</div>
          <div style="font-size:10px;color:var(--text3)">${lang==='ar'?'درجة ثقة المشكك':'Skeptic confidence score'}</div>
        </div>
      </div>`);
    }

    // Store skeptic assessment in memory
    if (agentMemory[ticker]) {
      agentMemory[ticker].skepticScore = confScore;
      agentMemory[ticker].skepticText = skepticText.slice(0, 300);
      scheduleMemorySave();
    }

    // ── STEP 7: Suggest next actions ──
    if (signal.aborted) return;
    const actions = [];
    if (peers.length >= 2) {
      const topPeer = peers[0];
      actions.push({ label: lang==='ar'?`📊 حلل ${topPeer}`:`📊 Analyse ${topPeer}`, onclick: `agentAnalysePeer('${topPeer}')` });
    }
    if (peers.length >= 2) {
      actions.push({ label: lang==='ar'?'🔄 مقارنة تفصيلية':'🔄 Deep compare', onclick: `agentDeepCompare('${ticker}')` });
    }
    if (priorTickers.length) {
      actions.push({ label: lang==='ar'?`↔️ قارن مع ${priorTickers[priorTickers.length-1]}`:`↔️ Compare with ${priorTickers[priorTickers.length-1]}`, onclick: `agentCrossCompare('${ticker}','${priorTickers[priorTickers.length-1]}')` });
    }
    actions.push({ label: lang==='ar'?'🔬 تحليل أعمق':'🔬 Deeper dive', onclick: `agentDeeperDive('${ticker}')` });

    showAgentActions(actions);
    setAgentStatus(false, lang === 'ar' ? `✓ اكتمل التحليل — ${Object.keys(agentMemory).length} شركة في الذاكرة` : `✓ Analysis complete — ${Object.keys(agentMemory).length} companies in memory`);

  } catch(err) {
    if (err.name === 'AbortError') return;
    addAgentStep('warning', '❌', lang==='ar'?'خطأ':'Error', `${err.message}\n\n${lang==='ar'?'تحقق من مفتاح API ورصيدك.':'Check your API key and balance.'}`);
    setAgentStatus(false, 'Error');
  } finally {
    if (regenBtn) regenBtn.disabled = false;
    aiAbortController = null;
  }
}

// ── AGENT FOLLOW-UP ACTIONS ──────────────────────────────────────────────────
async function agentAnalysePeer(ticker) {
  chatAnalyzeTicker(ticker);
}

async function agentDeepCompare(ticker) {
  if (!isAiReady() || !currentDash) return;
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const acts = document.getElementById('agentActions'); if (acts) acts.classList.add('hidden');
  setAgentStatus(true, lang==='ar'?'يجري مقارنة تفصيلية...':'Running deep comparison...');

  const peers = findPeers(ticker, 3);
  const allSnaps = [ticker, ...peers].map(t => buildTickerSnapshot(t)).filter(Boolean);
  const dataStr = allSnaps.map((s, i) => {
    const t = i === 0 ? ticker : peers[i-1];
    return `${t}: ${s.summary.split('\n').slice(0,3).join(', ')}`;
  }).join('\n\n');

  const sysP = lang === 'ar'
    ? 'أنت محلل مالي مقارن. قارن هذه الشركات في 3 فقرات: (1) من الأقوى ولماذا، (2) مفاجآت أو فجوات، (3) أي شركة تستحق المزيد من البحث. 150-200 كلمة. لا عناوين ولا نقاط.'
    : 'You are a comparative financial analyst. Compare these companies in 3 paragraphs: (1) which is strongest and why, (2) surprises or gaps, (3) which deserves further research. 150-200 words. No headers or bullets.';

  const el = addAgentStreamStep('comparison', '🔬', lang==='ar'?'مقارنة تفصيلية':'Deep Comparison');
  try {
    await streamWithRetry([{ role: 'system', content: sysP }, { role: 'user', content: dataStr }], el, signal);
    setAgentStatus(false, lang==='ar'?'✓ اكتملت المقارنة':'✓ Comparison complete');
  } catch(e) {
    if (e.name !== 'AbortError') addAgentStep('warning', '❌', 'Error', e.message);
    setAgentStatus(false, 'Error');
  }
  aiAbortController = null;
}

async function agentCrossCompare(tickerA, tickerB) {
  if (!isAiReady()) return;
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const acts = document.getElementById('agentActions'); if (acts) acts.classList.add('hidden');
  setAgentStatus(true, lang==='ar'?`يقارن ${tickerA} مع ${tickerB}...`:`Comparing ${tickerA} vs ${tickerB}...`);

  const snapA = buildTickerSnapshot(tickerA), snapB = buildTickerSnapshot(tickerB);
  if (!snapA || !snapB) { addAgentStep('warning', '⚠️', 'Error', 'Missing data'); setAgentStatus(false, 'Error'); aiAbortController = null; return; }

  const sysP = lang === 'ar'
    ? `أنت محلل مالي. قارن بين ${tickerA} و ${tickerB} في فقرتين: الفروقات الرئيسية ثم أيهما أقوى. 100-150 كلمة. لا عناوين.`
    : `You are a financial analyst. Compare ${tickerA} vs ${tickerB} in 2 paragraphs: key differences then which is stronger. 100-150 words. No headers.`;

  const el = addAgentStreamStep('comparison', '↔️', `${tickerA} vs ${tickerB}`);
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `${tickerA}:\n${snapA.summary}\n\n${tickerB}:\n${snapB.summary}` }
    ], el, signal);
    setAgentStatus(false, lang==='ar'?'✓ اكتملت المقارنة':'✓ Comparison complete');
  } catch(e) {
    if (e.name !== 'AbortError') addAgentStep('warning', '❌', 'Error', e.message);
    setAgentStatus(false, 'Error');
  }
  aiAbortController = null;
}

async function agentDeeperDive(ticker) {
  if (!isAiReady() || !currentDash) return;
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const acts = document.getElementById('agentActions'); if (acts) acts.classList.add('hidden');
  setAgentStatus(true, lang==='ar'?'يغوص أعمق...':'Diving deeper...');

  const snap = buildTickerSnapshot(ticker);
  if (!snap) { aiAbortController = null; return; }

  // Ask the AI to focus on what it finds most interesting
  const sysP = lang === 'ar'
    ? 'أنت محلل مالي متقدم. بناءً على البيانات، اختر أهم اتجاه واحد وقم بتحليله بعمق في فقرتين. مثلاً: تغير في الهوامش، نقطة انعطاف في النمو، إشارة من التدفق النقدي. 150-200 كلمة. لا عناوين.'
    : 'You are an advanced financial analyst. Based on the data, pick the single most interesting trend and analyse it deeply in 2 paragraphs. Examples: margin shift, growth inflection, cash flow signal. 150-200 words. No headers.';

  const priorAnalysis = agentMemory[ticker]?.analysis || '';
  const el = addAgentStreamStep('analysis', '🔬', lang==='ar'?'تحليل معمّق':'Deep Dive');
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `${snap.summary}\n\nPrior analysis summary: ${priorAnalysis.slice(0, 300)}...\n\nDig deeper — find something the prior analysis missed.` }
    ], el, signal);
    setAgentStatus(false, lang==='ar'?'✓ اكتمل التحليل المعمق':'✓ Deep dive complete');
    // Re-show actions
    const actions = [
      { label: lang==='ar'?'🔄 إعادة تحليل كامل':'🔄 Full re-analysis', onclick: `agentStart('${ticker}')` }
    ];
    const priors = Object.keys(agentMemory).filter(t => t !== ticker);
    if (priors.length) actions.push({ label: lang==='ar'?`↔️ قارن مع ${priors[priors.length-1]}`:`↔️ Compare with ${priors[priors.length-1]}`, onclick: `agentCrossCompare('${ticker}','${priors[priors.length-1]}')` });
    showAgentActions(actions);
  } catch(e) {
    if (e.name !== 'AbortError') addAgentStep('warning', '❌', 'Error', e.message);
    setAgentStatus(false, 'Error');
  }
  aiAbortController = null;
}

// Legacy wrapper removed — agent panel replaces static AI text

// ── AGENT CHAT (free-text follow-up) ─────────────────────────────────────────
async function agentChat(ticker) {
  const input = document.getElementById('agentChatInput');
  const sendBtn = document.getElementById('agentChatSend');
  if (!input || !input.value.trim()) return;
  if (!isAiReady()) { showPinModal(); return; }

  const question = input.value.trim();
  input.value = '';
  if (sendBtn) sendBtn.disabled = true;

  // Show user's question
  addAgentStep('observation', '💬', lang==='ar'?'سؤالك':'Your question', question);

  // Cancel previous
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  setAgentStatus(true, lang==='ar'?'يفكر في إجابتك...':'Thinking...');

  const snap = buildTickerSnapshot(ticker);
  const priorContext = agentMemory[ticker]?.analysis?.slice(0, 500) || '';
  const newsCtx = agentMemory[ticker]?.news ? `\nRecent news:\n${agentMemory[ticker].news}` : '';
  const bench = getSectorBenchmarks(ticker);
  let sectorCtx = '';
  if (bench) {
    sectorCtx = `\nSector benchmarks (${bench.sector}, ${bench.count} companies): Avg score ${bench.avgOverall?.toFixed(1)}, Avg growth ${bench.avgGrowth?.toFixed(1)}%, Avg margin ${bench.avgMargin?.toFixed(1)}%, Avg D/E ${bench.avgDTE?.toFixed(2)}x.`;
  }

  const sysP = lang === 'ar'
    ? 'أنت وكيل مالي ذكي. أجب على السؤال بدقة مستخدماً جميع البيانات المتاحة بما في ذلك الأخبار والتقديرات. فقرة أو اثنتين. لا عناوين.'
    : 'You are a financial AI agent. Answer precisely using ALL provided data including news headlines, analyst estimates, and sector benchmarks. 1-2 paragraphs. No headers.';

  const el = addAgentStreamStep('analysis', '🤖', lang==='ar'?'إجابة الوكيل':'Agent response');
  try {
    const response = await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `Company: ${ticker}\n${snap?.summary || 'No data'}\n${sectorCtx}${newsCtx}\n\nPrior analysis: ${priorContext}\n\nUser question: ${question}` }
    ], el, signal);

    // Auto-detect if we should generate a visual
    const lq = question.toLowerCase();
    if (lq.includes('chart') || lq.includes('graph') || lq.includes('visual') || lq.includes('compare') || lq.includes('رسم') || lq.includes('قارن') || lq.includes('مخطط')) {
      await agentGenerateVisual(ticker, question, signal);
    }

    setAgentStatus(false, lang==='ar'?'✓ تم الرد':'✓ Answered');
  } catch(e) {
    if (e.name !== 'AbortError') addAgentStep('warning', '❌', 'Error', e.message);
    setAgentStatus(false, 'Error');
  }
  if (sendBtn) sendBtn.disabled = false;
  aiAbortController = null;
}

// ── AGENT INLINE VISUALS ─────────────────────────────────────────────────────
async function agentGenerateVisual(ticker, context, signal) {
  if (!currentDash) return;
  const snap = buildTickerSnapshot(ticker);
  if (!snap) return;

  const conv = agentConv(); if (!conv) return;

  // Determine what visual to show based on context
  const lc = context.toLowerCase();
  const peers = findPeers(ticker, 4);

  if ((lc.includes('compare') || lc.includes('peer') || lc.includes('قارن') || lc.includes('أقران')) && peers.length) {
    // Score comparison bar chart
    const allTickers = [ticker, ...peers.slice(0, 4)];
    const allData = allTickers.map(t => {
      const s = buildTickerSnapshot(t);
      return { ticker: t, score: s?.scores.overall || 0 };
    });
    const maxScore = 10;
    let html = `<div class="agent-mini-chart"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">${lang==='ar'?'مقارنة الدرجات':'Score Comparison'}</div><div class="agent-bar-chart">`;
    allData.forEach(d => {
      const pct = (d.score / maxScore * 100);
      const col = sCol(d.score);
      const isCurrent = d.ticker === ticker;
      html += `<div class="agent-bar">
        <div class="agent-bar-val" style="color:${col}">${d.score}</div>
        <div class="agent-bar-fill" style="height:${pct}%;background:${col}${isCurrent?'':'66'};${isCurrent?'border:2px solid '+col:''}"></div>
        <div class="agent-bar-lbl" style="${isCurrent?'font-weight:700;color:var(--text)':''}">${d.ticker}</div>
      </div>`;
    });
    html += '</div></div>';
    const div = document.createElement('div');
    div.className = 'agent-step observation';
    div.innerHTML = `<div class="step-hdr"><span class="step-icon">📊</span><span class="step-label observation">${lang==='ar'?'رسم بياني':'Visual'}</span></div>${html}`;
    conv.appendChild(div);
    conv.scrollTop = conv.scrollHeight;
  } else if (lc.includes('margin') || lc.includes('هامش') || lc.includes('profit') || lc.includes('ربح')) {
    // Margin trend
    const mData = snap.m.prof;
    let html = `<div class="agent-mini-chart"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">${lang==='ar'?'اتجاه الهوامش':'Margin Trend'}</div><div class="agent-bar-chart">`;
    mData.forEach(p => {
      const val = p.net_margin;
      if (val == null) return;
      const pct = Math.max(5, Math.min(100, Math.abs(val) * 2));
      const col = val >= 0 ? 'var(--green)' : 'var(--red)';
      html += `<div class="agent-bar">
        <div class="agent-bar-val" style="color:${col}">${val.toFixed(1)}%</div>
        <div class="agent-bar-fill" style="height:${pct}%;background:${col}88"></div>
        <div class="agent-bar-lbl">${p.year}</div>
      </div>`;
    });
    html += '</div></div>';
    const div = document.createElement('div');
    div.className = 'agent-step observation';
    div.innerHTML = `<div class="step-hdr"><span class="step-icon">📊</span><span class="step-label observation">${lang==='ar'?'رسم بياني':'Visual'}</span></div>${html}`;
    conv.appendChild(div);
    conv.scrollTop = conv.scrollHeight;
  } else if (lc.includes('growth') || lc.includes('نمو') || lc.includes('revenue') || lc.includes('إيرادات')) {
    // Growth trend
    const yData = snap.m.yoy;
    let html = `<div class="agent-mini-chart"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">${lang==='ar'?'اتجاه النمو':'Growth Trend'}</div><div class="agent-bar-chart">`;
    yData.forEach(y => {
      const val = y.revenue_growth;
      if (val == null) return;
      const pct = Math.max(5, Math.min(100, Math.abs(val) * 2));
      const col = val >= 0 ? 'var(--green)' : 'var(--red)';
      html += `<div class="agent-bar">
        <div class="agent-bar-val" style="color:${col}">${val>0?'+':''}${val.toFixed(1)}%</div>
        <div class="agent-bar-fill" style="height:${pct}%;background:${col}88"></div>
        <div class="agent-bar-lbl">${y.year}</div>
      </div>`;
    });
    html += '</div></div>';
    const div = document.createElement('div');
    div.className = 'agent-step observation';
    div.innerHTML = `<div class="step-hdr"><span class="step-icon">📊</span><span class="step-label observation">${lang==='ar'?'رسم بياني':'Visual'}</span></div>${html}`;
    conv.appendChild(div);
    conv.scrollTop = conv.scrollHeight;
  }
}

// ── CHAT-FIRST INTERFACE ────────────────────────────────────────────────────
let chatHistory = [];

function initChatView() {
  const ws = document.getElementById('welcomeScreen');
  if (ws) ws.classList.add('hidden');
  const cv = document.getElementById('chatView');
  if (cv) cv.classList.remove('hidden');

  const nTickers = TICKERS.length;
  const nSectors = new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean)).size;
  const isAr = lang === 'ar';

  const h2 = document.getElementById('chatWelcomeH2');
  if (h2) h2.textContent = isAr ? 'ماذا تريد أن تبحث اليوم؟' : 'What would you like to research?';
  const sub = document.getElementById('chatWelcomeP');
  if (sub) sub.textContent = isAr
    ? `${nTickers} شركة في ${nSectors} قطاعات جاهزة للتحليل`
    : `${nTickers} companies across ${nSectors} sectors ready for analysis`;

  const knownTickers = ['AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL', 'META', 'TSLA'];
  const featuredTicker = knownTickers.find(t => TICKERS.includes(t)) || TICKERS[0] || 'AAPL';

  const suggestions = [
    { label: isAr ? `📊 حلل ${featuredTicker}` : `📊 Analyze ${featuredTicker}`, query: `Analyze ${featuredTicker}` },
    { label: isAr ? '↔️ قارن MSFT مع GOOGL' : '↔️ Compare MSFT vs GOOGL', query: 'Compare MSFT vs GOOGL' },
    { label: isAr ? '🔍 أسهم مقومة بأقل من قيمتها' : '🔍 Undervalued stocks', query: 'Show me undervalued stocks' },
    { label: isAr ? '📈 أقوى القطاعات' : '📈 Strongest sectors', query: 'Which sectors are strongest?' },
    { label: isAr ? '🏆 أفضل 5 شركات' : '🏆 Top 5 companies', query: 'Top 5 highest scoring companies' },
  ];

  const sugEl = document.getElementById('chatSuggestions');
  if (sugEl) sugEl.innerHTML = suggestions.map(s =>
    `<button class="chat-chip" onclick="handleChatChip('${s.query.replace(/'/g, "\\'")}')">${s.label}</button>`
  ).join('');

  document.getElementById('hdrSearch').disabled = false;
  document.getElementById('hdrSearch').placeholder = isAr ? 'ابحث عن رمز أو شركة...' : 'Search ticker or company...';
}

function handleChatChip(query) {
  const input = document.getElementById('chatInput');
  if (input) input.value = query;
  handleChatInput();
}

function addChatMessage(role, content, opts = {}) {
  const welcome = document.getElementById('chatWelcome');
  if (welcome) welcome.style.display = 'none';

  const container = document.getElementById('chatMessages');
  if (!container) return null;

  const msg = document.createElement('div');
  msg.className = `chat-msg ${role}`;

  const bubble = document.createElement('div');
  bubble.className = 'chat-msg-bubble';

  if (role === 'user') {
    bubble.textContent = content;
  } else {
    if (content) bubble.innerHTML = content;
  }

  msg.appendChild(bubble);

  if (opts.richHtml) {
    const rich = document.createElement('div');
    rich.className = 'chat-rich-card';
    rich.innerHTML = opts.richHtml;
    bubble.appendChild(rich);
  }

  if (opts.actions) {
    const actDiv = document.createElement('div');
    actDiv.className = 'chat-actions';
    actDiv.innerHTML = opts.actions.map(a =>
      `<button class="chat-action-btn${a.primary ? ' primary' : ''}" onclick="${a.onclick}">${a.label}</button>`
    ).join('');
    bubble.appendChild(actDiv);
  }

  container.appendChild(msg);
  container.scrollTop = container.scrollHeight;

  chatHistory.push({ role, content: content || '', ticker: opts.ticker || null, timestamp: Date.now() });
  return bubble;
}

function addChatStreamBubble() {
  const welcome = document.getElementById('chatWelcome');
  if (welcome) welcome.style.display = 'none';
  const container = document.getElementById('chatMessages');
  if (!container) return null;

  const msg = document.createElement('div');
  msg.className = 'chat-msg agent';
  const bubble = document.createElement('div');
  bubble.className = 'chat-msg-bubble';
  msg.appendChild(bubble);
  container.appendChild(msg);
  container.scrollTop = container.scrollHeight;
  return bubble;
}

function addChatTyping() {
  const container = document.getElementById('chatMessages');
  if (!container) return null;
  const msg = document.createElement('div');
  msg.className = 'chat-msg agent';
  msg.id = 'chatTypingIndicator';
  msg.innerHTML = '<div class="chat-msg-bubble"><div class="chat-typing"><div class="chat-typing-dot"></div><div class="chat-typing-dot"></div><div class="chat-typing-dot"></div></div></div>';
  container.appendChild(msg);
  container.scrollTop = container.scrollHeight;
  return msg;
}

function removeChatTyping() {
  const el = document.getElementById('chatTypingIndicator');
  if (el) el.remove();
}

function scrollChatToBottom() {
  const container = document.getElementById('chatMessages');
  if (container) container.scrollTop = container.scrollHeight;
}

function handleChatInput() {
  const input = document.getElementById('chatInput');
  if (!input || !input.value.trim()) return;
  const text = input.value.trim();
  input.value = '';
  addChatMessage('user', text);
  routeChatMessage(text);
}

function routeChatMessage(text) {
  const lower = text.toLowerCase();
  const allTickers = detectAllTickers(text);

  // ── ARTICLE-CHAT MODE: if discussing a news article, answer from it ──
  // Exit if the user clearly wants to leave (types a bare ticker or an explicit command).
  if (_activeArticleContext) {
    const wantsOut = /^(exit|stop|done|back|clear|new topic)\b/i.test(text.trim())
      || (TICKERS.includes(text.trim().toUpperCase()) && text.trim().length <= 5);
    if (wantsOut) {
      clearArticleChat();
      // fall through to normal routing
    } else {
      answerAboutArticle(text);
      return;
    }
  }

  // ── FAST-PATH SHORTCUTS: only the most obvious direct commands ──

  // 1. Bare ticker symbol(s) — "AAPL" or "NVDA"
  const trimmed = text.trim().toUpperCase();
  if (TICKERS.includes(trimmed) && trimmed.length <= 5) {
    chatAnalyzeTicker(trimmed); return;
  }

  // 2. Explicit ticker comparison — "AAPL vs MSFT", "compare AAPL and MSFT"
  const cmpMatch = lower.match(/^\s*(?:compare\s+)?(\w+)\s+(?:vs\.?|versus|against)\s+(\w+)\s*\??$/i) ||
                   lower.match(/^\s*compare\s+(\w+)\s+(?:to|with|and)\s+(\w+)\s*\??$/i);
  if (cmpMatch) {
    const a = cmpMatch[1].toUpperCase(), b = cmpMatch[2].toUpperCase();
    if (TICKERS.includes(a) && TICKERS.includes(b)) { chatCompareTickers(a, b); return; }
  }

  // 3. Heatmap / watchlist / portfolio — explicit UI commands

  if (lower.match(/^(show |open |display )?heatmap\s*$|^خريطة$/i)) { showHeatmap(); return; }

  // "show portfolio" / "open portfolio" / "view portfolio" → open the dashboard view
  if (lower.match(/^(show|open|view|display)\s+(my\s+)?(portfolio|holdings)\s*$/i)) {
    if (portfolio?.length) { showPortfolioDashboard(); return; }
    addChatMessage('agent', lang==='ar' ? '💼 المحفظة فارغة. أضف مراكزك من اللوحة الجانبية.' : '💼 Portfolio is empty. Add positions from the sidebar.');
    return;
  }

  // "analyze/review my portfolio" / "how is my portfolio" / "portfolio analysis" → council analysis
  if (lower.match(/^(analy[sz]e|review|check)\s+(my\s+)?(portfolio|holdings)\s*$/i)
   || lower.match(/^(how['']?s?|how is)\s+(my\s+)?(portfolio|holdings)\s*(doing|going|looking)?\s*\??$/i)
   || lower.match(/^(my\s+)?(portfolio|holdings)\s+(analysis|review|status|risk|breakdown)\s*\??$/i)) {
    if (portfolio?.length) { askCouncilAboutPortfolio(); return; }
    addChatMessage('agent', lang==='ar' ? '💼 المحفظة فارغة. أضف مراكزك من اللوحة الجانبية لتحليلها.' : '💼 Your portfolio is empty. Add positions from the sidebar to analyze.');
    return;
  }

  // Watchlist (kept for backwards compat — older terminology)
  if (lower.match(/^(show |analyse |analyze )?(my )?watchlist\s*$/i)) { chatWatchlistAnalysis(); return; }

  // "recommend stocks" / "what should I buy" / "give me recommendations" → portfolio advisor
  if (lower.match(/^(recommend|suggest)\s+(some\s+)?(stocks?|positions?|picks?|investments?)/i)
   || lower.match(/^(what\s+should\s+i\s+(buy|invest)|give\s+me\s+(some\s+)?(recommendations?|picks?|suggestions?))/i)
   || lower.match(/^get\s+(some\s+)?recommendations?/i)) {
    startPortfolioAdvisor(); return;
  }

  // ── EVERYTHING ELSE → AGENT (AI decides what tools to use) ──
  runAgent(text);
}

function detectTicker(text) {
  const words = text.toUpperCase().replace(/[^A-Z0-9\s]/g, '').split(/\s+/);
  const common = new Set(['A','ALL','ARE','IT','ON','OR','AN','SO','DO','GO','HAS','NOW','OUT','KEY','WELL','FAST','TECH','PAY','MAN','LOW','BIG','RE','AI']);
  for (const w of words) {
    if (TICKERS.includes(w) && !common.has(w)) return w;
  }
  for (const w of words) {
    if (TICKERS.includes(w) && common.has(w) && words.length <= 2) return w;
  }
  return null;
}

function detectAllTickers(text) {
  const words = text.toUpperCase().replace(/[^A-Z0-9\s]/g, '').split(/\s+/);
  const common = new Set(['A','ALL','ARE','IT','ON','OR','AN','SO','DO','GO','HAS','NOW','OUT','KEY','WELL','FAST','TECH','PAY','MAN','LOW','BIG','RE','AI']);
  const found = [];
  for (const w of words) {
    if (TICKERS.includes(w) && !common.has(w) && !found.includes(w)) found.push(w);
  }
  return found;
}

// ════════════════════════════════════════════════════════════════════════════
// AGENT LOOP — tool-using AI agent
// ════════════════════════════════════════════════════════════════════════════

const AGENT_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'screen_by_metric',
      description: 'Rank stocks by a financial metric. Returns top/bottom N tickers with values. Use for "top 5 by market cap", "highest revenue companies", "lowest P/E ratios", etc.',
      parameters: {
        type: 'object',
        properties: {
          metric: { type: 'string', enum: ['market_cap','revenue','pe','eps','roe','net_income','overall_score','growth_score','profit_score','dividend_yield','net_margin'], description: 'Which metric to rank by' },
          direction: { type: 'string', enum: ['highest','lowest'], description: 'Sort direction' },
          top_n: { type: 'number', description: 'How many results to return (max 30)' },
          sector: { type: 'string', description: 'Optional sector filter (e.g., "Technology")' }
        },
        required: ['metric','direction','top_n']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_ticker_snapshot',
      description: 'Get full financial snapshot for one ticker: 5yr financials, scores, peer comparison, sector rank.',
      parameters: {
        type: 'object',
        properties: { ticker: { type: 'string', description: 'Stock ticker symbol' } },
        required: ['ticker']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'compare_tickers',
      description: 'Side-by-side comparison of 2-5 tickers (scores, key metrics, sector rank).',
      parameters: {
        type: 'object',
        properties: { tickers: { type: 'array', items: { type: 'string' }, description: '2-5 ticker symbols to compare' } },
        required: ['tickers']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_sector_overview',
      description: 'Get overview of a sector: stock count, total market cap, avg score, top members.',
      parameters: {
        type: 'object',
        properties: { sector: { type: 'string', description: 'Sector name (e.g., "Technology", "Energy")' } },
        required: ['sector']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'screen_by_filter',
      description: 'Filter stocks by qualitative criteria (undervalued, growth, dividend, etc.). Returns matching tickers.',
      parameters: {
        type: 'object',
        properties: {
          filter: { type: 'string', enum: ['undervalued','growth','dividend','low_debt','high_margin','large_cap','small_cap','cash_rich'], description: 'Filter type' },
          sector: { type: 'string', description: 'Optional sector filter' },
          top_n: { type: 'number', description: 'Max results (default 10)' }
        },
        required: ['filter']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_universe_stats',
      description: 'Get overview of the entire universe: total stocks, sector breakdown, top names. Use this when the user asks broad questions about the market.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'screen_by_performance',
      description: 'Find stocks by ACTUAL price performance over any period. Use for "stocks that went up over the last year", "best/worst performers this month", "stocks that gained 10%+ in past 7 days". Fetches real historical prices.',
      parameters: {
        type: 'object',
        properties: {
          period_days: { type: 'number', description: 'Lookback period in calendar days. 1=today, 7=week, 30=month, 90=quarter, 365=year.' },
          direction: { type: 'string', enum: ['top','bottom'], description: 'top = biggest gainers, bottom = biggest losers' },
          top_n: { type: 'number', description: 'How many results (max 20)' },
          sector: { type: 'string', description: 'Optional sector filter (e.g., "Technology"). Strongly recommended for performance queries to keep API calls fast.' },
          min_change_pct: { type: 'number', description: 'Optional absolute % threshold (e.g., 10 means returns only stocks that moved at least ±10%)' }
        },
        required: ['period_days','direction','top_n']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'recommend_for_profile',
      description: 'Get a ranked list of stock candidates suitable for an investor profile (conservative/moderate/aggressive). Returns candidates filtered by quality score, valuation, market cap, leverage, and dividend criteria for the risk level. The top 8 are ENRICHED with: recent news headlines, upcoming earnings date (with days_until), analyst consensus (rating + price target + upside %), and a knowledge-graph "kg_overlap_with_portfolio" check that warns if recommending this ticker would deepen existing thematic exposure. Also returns portfolio_thematic_exposure showing what themes the user is already clustered around. Use this for ALL recommendation requests.',
      parameters: {
        type: 'object',
        properties: {
          risk: { type: 'string', enum: ['conservative','moderate','aggressive'], description: 'Investor risk tolerance' },
          exclude_sectors: { type: 'array', items: { type: 'string' }, description: 'Sectors to exclude from candidates' },
          exclude_owned: { type: 'boolean', description: 'Skip tickers already in user portfolio (recommend NEW positions). Default true.' }
        },
        required: ['risk']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_portfolio',
      description: 'Get the user\'s current portfolio: list of held positions (ticker, shares, cost basis, current value, P&L), sector allocation, weighted quality score, day/total P&L. Use this whenever the user asks about "my portfolio", "my holdings", "how am I doing", concentration risk, or wants analysis specific to what they own.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_related_companies',
      description: 'Find OTHER companies related to a given ticker for causal-chain analysis. Returns competitors, suppliers, customers, partners, co-dependent companies, AND thematic peers (companies riding the same secular trend like AI/EV/cloud) with evidence and strength scores. Use this when the user asks "who else gets hit if X happens?", "who benefits from X trend?", "what are X\'s competitors?", or for any analysis needing second-order effects across supply chain, competitive landscape, or shared tailwinds.',
      parameters: {
        type: 'object',
        properties: {
          ticker: { type: 'string', description: 'Source ticker to find relationships for' },
          type: { type: 'string', enum: ['competitor','supplier','customer','partner','co_dependent','thematic_peer','all'], description: 'Filter by relationship type. Default: all' }
        },
        required: ['ticker']
      }
    }
  }
];

// Tool executor — maps function name to actual implementation
async function executeAgentTool(name, args) {
  try {
    switch (name) {
      case 'screen_by_metric': return tool_screenByMetric(args);
      case 'get_ticker_snapshot': return tool_getTickerSnapshot(args);
      case 'compare_tickers': return tool_compareTickers(args);
      case 'get_sector_overview': return tool_getSectorOverview(args);
      case 'screen_by_filter': return tool_screenByFilter(args);
      case 'get_universe_stats': return tool_getUniverseStats();
      case 'screen_by_performance': return await tool_screenByPerformance(args);
      case 'get_related_companies': return await tool_getRelatedCompanies(args);
      case 'get_portfolio': return tool_getPortfolio();
      case 'recommend_for_profile': return await tool_recommendForProfile(args);
      default: return { error: `Unknown tool: ${name}` };
    }
  } catch (e) {
    return { error: e.message };
  }
}

function tool_screenByMetric({ metric, direction, top_n, sector }) {
  const n = Math.min(top_n || 10, 30);
  let pool = TICKERS.filter(t => ANNUAL[t]?.length);
  if (sector) pool = pool.filter(t => STOCK[t]?.sector?.toLowerCase() === sector.toLowerCase());
  const getValue = (t) => {
    const m = calcMetrics(ANNUAL[t]);
    const s = calcScores(m);
    const stk = STOCK[t] || {};
    switch (metric) {
      case 'market_cap': return stk.marketCap;
      case 'revenue': return m.latest?.revenue;
      case 'pe': return stk.pe > 0 ? stk.pe : null;
      case 'eps': return m.latest?.eps_diluted;
      case 'roe': return m.prof?.[m.prof.length-1]?.roe;
      case 'net_income': return m.latest?.net_income;
      case 'overall_score': return s.overall;
      case 'growth_score': return s.growth;
      case 'profit_score': return s.profitability;
      case 'net_margin': return m.prof?.[m.prof.length-1]?.net_margin;
      case 'dividend_yield': return stk.dividend && stk.price ? (stk.dividend / stk.price * 100) : null;
      default: return null;
    }
  };
  const data = pool.map(t => ({ ticker: t, value: getValue(t), sector: STOCK[t]?.sector })).filter(x => x.value != null);
  data.sort((a, b) => direction === 'lowest' ? a.value - b.value : b.value - a.value);
  return { metric, direction, count: data.length, top_results: data.slice(0, n) };
}

function priceContext(stk) {
  if (!stk?.price) return null;
  const ctx = { price: parseFloat(stk.price.toFixed(2)) };
  const lo = stk.fiftyTwoWeekLow ?? stk.yearLow;
  const hi = stk.fiftyTwoWeekHigh ?? stk.yearHigh;
  if (lo && hi && hi > lo) {
    const pctOfRange = ((stk.price - lo) / (hi - lo) * 100);
    const pctFromHigh = ((stk.price - hi) / hi * 100);
    ctx.year_low = parseFloat(lo.toFixed(2));
    ctx.year_high = parseFloat(hi.toFixed(2));
    ctx.position_in_52w_range_pct = parseInt(pctOfRange.toFixed(0));
    ctx.pct_from_52w_high = parseFloat(pctFromHigh.toFixed(1));
  }
  if (stk.changePct != null) ctx.change_today_pct = parseFloat(stk.changePct.toFixed(2));
  return ctx;
}

function tool_getTickerSnapshot({ ticker }) {
  const t = ticker.toUpperCase();
  if (!TICKERS.includes(t)) return { error: `Ticker ${t} not found` };
  const snap = buildTickerSnapshot(t);
  if (!snap) return { error: `No data for ${t}` };
  const peers = findPeers(t, 5);
  const sectorPeers = TICKERS.filter(x => STOCK[x]?.sector === snap.stk.sector && ANNUAL[x]?.length);
  const sectorScores = sectorPeers.map(x => ({ x, s: calcScores(calcMetrics(ANNUAL[x])).overall })).sort((a,b) => b.s - a.s);
  const rank = sectorScores.findIndex(x => x.x === t) + 1;
  return { ticker: t, summary: snap.summary, price: priceContext(snap.stk), peers, sector_rank: `${rank} of ${sectorPeers.length}` };
}

function tool_compareTickers({ tickers }) {
  const results = tickers.slice(0, 5).map(tk => {
    const t = tk.toUpperCase();
    if (!TICKERS.includes(t) || !ANNUAL[t]?.length) return { ticker: t, error: 'No data' };
    const snap = buildTickerSnapshot(t);
    return snap ? {
      ticker: t,
      scores: snap.scores,
      price: priceContext(snap.stk),
      market_cap: snap.stk.marketCap,
      pe: snap.stk.pe,
      sector: snap.stk.sector,
      latest_revenue: snap.m.latest?.revenue,
      latest_net_income: snap.m.latest?.net_income
    } : { ticker: t, error: 'Snapshot failed' };
  });
  return { comparisons: results };
}

function tool_getSectorOverview({ sector }) {
  const members = TICKERS.filter(t => STOCK[t]?.sector?.toLowerCase() === sector.toLowerCase() && ANNUAL[t]?.length);
  if (!members.length) return { error: `No stocks found in sector "${sector}"` };
  const totalMcap = members.reduce((s, t) => s + (STOCK[t]?.marketCap || 0), 0);
  const scored = members.map(t => ({ t, score: calcScores(calcMetrics(ANNUAL[t])).overall, mcap: STOCK[t]?.marketCap || 0 })).sort((a,b) => b.mcap - a.mcap);
  const avgScore = (scored.reduce((s, x) => s + x.score, 0) / scored.length).toFixed(1);
  return {
    sector, count: members.length, total_market_cap: fM(totalMcap), avg_score: avgScore,
    top_15_by_mcap: scored.slice(0, 15).map(x => ({ ticker: x.t, mcap: fM(x.mcap), score: x.score }))
  };
}

function tool_screenByFilter({ filter, sector, top_n }) {
  const n = Math.min(top_n || 10, 30);
  let pool = TICKERS.filter(t => ANNUAL[t]?.length);
  if (sector) pool = pool.filter(t => STOCK[t]?.sector?.toLowerCase() === sector.toLowerCase());
  switch (filter) {
    case 'undervalued': pool = pool.filter(t => STOCK[t]?.pe > 0 && STOCK[t].pe < 15); break;
    case 'growth': pool = pool.filter(t => { const g = calcMetrics(ANNUAL[t]).yoy?.slice(-1)[0]?.revenue_growth; return g && g > 15; }); break;
    case 'dividend': pool = pool.filter(t => STOCK[t]?.dividend > 0); break;
    case 'low_debt': pool = pool.filter(t => { const d = calcMetrics(ANNUAL[t]).lev?.slice(-1)[0]?.dte; return d != null && d < 0.5; }); break;
    case 'high_margin': pool = pool.filter(t => { const nm = calcMetrics(ANNUAL[t]).prof?.slice(-1)[0]?.net_margin; return nm > 20; }); break;
    case 'large_cap': pool = pool.filter(t => STOCK[t]?.marketCap > 100e9); break;
    case 'small_cap': pool = pool.filter(t => STOCK[t]?.marketCap && STOCK[t].marketCap < 10e9); break;
    case 'cash_rich': pool = pool.filter(t => calcMetrics(ANNUAL[t]).latest?.free_cash_flow > 0); break;
  }
  const ranked = pool.map(t => ({ ticker: t, score: calcScores(calcMetrics(ANNUAL[t])).overall, mcap: STOCK[t]?.marketCap || 0, sector: STOCK[t]?.sector })).sort((a,b) => b.score - a.score);
  return { filter, count: ranked.length, top_results: ranked.slice(0, n) };
}

function tool_getUniverseStats() {
  return { full_universe: buildUniverseDigest() };
}

// In-memory cache for performance results within a session (avoid duplicate FMP calls)
const _perfCache = new Map();

async function tool_screenByPerformance({ period_days, direction, top_n, sector, min_change_pct }) {
  if (!FMP_QUOTE_KEY) return { error: 'Price data API not configured' };
  const days = Math.max(1, Math.min(period_days || 30, 365 * 5));
  const n = Math.min(top_n || 10, 20);

  // Build candidate pool: filter by sector if given, otherwise cap at top-100 by market cap
  let pool = TICKERS.filter(t => ANNUAL[t]?.length && STOCK[t]?.marketCap);
  if (sector) pool = pool.filter(t => STOCK[t]?.sector?.toLowerCase() === sector.toLowerCase());
  if (!sector) pool = pool.sort((a, b) => (STOCK[b]?.marketCap || 0) - (STOCK[a]?.marketCap || 0)).slice(0, 100);

  // 1-day uses cached changePct (free, instant)
  let perfData;
  if (days === 1) {
    perfData = pool.filter(t => STOCK[t]?.changePct != null).map(t => ({
      ticker: t, perf: STOCK[t].changePct, price: STOCK[t].price || 0, sector: STOCK[t].sector
    }));
  } else {
    const cacheKey = `${days}d:${pool.join(',')}`;
    if (_perfCache.has(cacheKey) && Date.now() - _perfCache.get(cacheKey).t < 60 * 60 * 1000) {
      perfData = _perfCache.get(cacheKey).data;
    } else {
      const lookback = Math.ceil(days * 1.4) + 3;
      const fromDate = new Date(); fromDate.setDate(fromDate.getDate() - lookback);
      const fromStr = fromDate.toISOString().slice(0, 10);
      const targetBars = Math.max(2, Math.round(days * 5 / 7));

      perfData = [];
      const BATCH = 5;
      for (let i = 0; i < pool.length; i += BATCH) {
        const batch = pool.slice(i, i + BATCH);
        const results = await Promise.all(batch.map(async t => {
          try {
            const cacheT = `t:${t}:${days}`;
            if (_perfCache.has(cacheT) && Date.now() - _perfCache.get(cacheT).t < 60 * 60 * 1000) {
              return _perfCache.get(cacheT).data;
            }
            const res = await fetch(`https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${t}&from=${fromStr}&apikey=${FMP_QUOTE_KEY}`);
            if (!res.ok) return null;
            const data = await res.json();
            if (!Array.isArray(data) || data.length < 2) return null;
            const newest = data[0]?.close || data[0]?.price;
            const oldestIdx = Math.min(targetBars - 1, data.length - 1);
            const oldest = data[oldestIdx]?.close || data[oldestIdx]?.price;
            if (!oldest || !newest) return null;
            const item = { ticker: t, perf: ((newest - oldest) / oldest * 100), price: newest, sector: STOCK[t]?.sector };
            _perfCache.set(cacheT, { t: Date.now(), data: item });
            return item;
          } catch { return null; }
        }));
        perfData.push(...results.filter(Boolean));
        if (i + BATCH < pool.length) await new Promise(r => setTimeout(r, 250));
      }
      _perfCache.set(cacheKey, { t: Date.now(), data: perfData });
    }
  }

  // Apply threshold
  if (min_change_pct != null) {
    const abs = Math.abs(min_change_pct);
    perfData = perfData.filter(d => Math.abs(d.perf) >= abs);
  }

  perfData.sort((a, b) => direction === 'bottom' ? a.perf - b.perf : b.perf - a.perf);
  const results = perfData.slice(0, n);

  return {
    period_days: days,
    period_label: days === 1 ? '1 day' : days <= 7 ? `${days} days` : days <= 90 ? `${Math.round(days/7)} weeks` : `${Math.round(days/30)} months`,
    direction,
    sector: sector || 'all (top 100 by mcap)',
    count: results.length,
    universe_size: pool.length,
    results: results.map(r => ({ ticker: r.ticker, sector: r.sector, performance_pct: parseFloat(r.perf.toFixed(2)), price: parseFloat(r.price.toFixed(2)) }))
  };
}

// ── PORTFOLIO (Commit A) ────────────────────────────────────────────────────
// Computes aggregate metrics from `portfolio` array of {ticker, shares, costBasis?}.
function calcPortfolio() {
  if (!portfolio.length) return null;
  const positions = portfolio.map(p => {
    const stk = STOCK[p.ticker] || {};
    const price = stk.price || 0;
    const value = price * p.shares;
    const cost = p.costBasis != null ? p.costBasis * p.shares : null;
    const unrealized = cost != null ? value - cost : null;
    const unrealizedPct = cost ? (unrealized / cost) * 100 : null;
    const dayPct = stk.changePct != null ? stk.changePct : null;
    const dayPL = dayPct != null && value ? value * (dayPct / 100) : null;
    const m = ANNUAL[p.ticker]?.length ? calcMetrics(ANNUAL[p.ticker]) : null;
    const score = m ? calcScores(m).overall : null;
    return {
      ticker: p.ticker,
      shares: p.shares,
      costBasis: p.costBasis,
      price, value, cost, unrealized, unrealizedPct,
      dayPct, dayPL, score,
      sector: stk.sector || 'Unknown'
    };
  });
  const totalValue = positions.reduce((s, p) => s + (p.value || 0), 0);
  const totalCost = positions.reduce((s, p) => s + (p.cost || 0), 0);
  const totalUnrealized = totalCost ? totalValue - totalCost : null;
  const totalUnrealizedPct = totalCost ? (totalUnrealized / totalCost) * 100 : null;
  const totalDayPL = positions.reduce((s, p) => s + (p.dayPL || 0), 0);
  const totalDayPct = totalValue ? (totalDayPL / totalValue) * 100 : null;

  // Weighted overall score (by value)
  const scored = positions.filter(p => p.score != null && p.value > 0);
  const weightedScore = scored.length
    ? scored.reduce((s, p) => s + p.score * p.value, 0) / scored.reduce((s, p) => s + p.value, 0)
    : null;

  // Sector allocation
  const sectorMap = {};
  positions.forEach(p => { sectorMap[p.sector] = (sectorMap[p.sector] || 0) + p.value; });
  const sectorAlloc = Object.entries(sectorMap)
    .map(([sector, v]) => ({ sector, value: v, pct: totalValue ? (v / totalValue) * 100 : 0 }))
    .sort((a, b) => b.value - a.value);

  return {
    positions: positions.sort((a, b) => b.value - a.value),
    totalValue, totalCost, totalUnrealized, totalUnrealizedPct,
    totalDayPL, totalDayPct, weightedScore, sectorAlloc, count: positions.length
  };
}

// Parse CSV input like "AAPL, 100, 150.50" → array of position objects.
// Accepts: ticker / ticker,shares / ticker,shares,costBasis per line.
function parsePortfolioCSV(text) {
  if (!text) return { positions: [], errors: [] };
  const positions = [];
  const errors = [];
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  for (const line of lines) {
    const parts = line.split(/[,\s]+/).filter(Boolean);
    if (!parts.length) continue;
    const ticker = parts[0].toUpperCase();
    if (!TICKERS.includes(ticker)) { errors.push(`${ticker}: not in universe`); continue; }
    const shares = parts[1] ? parseFloat(parts[1]) : null;
    if (shares == null || !isFinite(shares) || shares <= 0) { errors.push(`${ticker}: invalid shares`); continue; }
    const costBasis = parts[2] ? parseFloat(parts[2]) : null;
    positions.push({
      ticker, shares,
      costBasis: (costBasis != null && isFinite(costBasis) && costBasis > 0) ? costBasis : null,
      addedAt: new Date().toISOString()
    });
  }
  return { positions, errors };
}

// Add or update positions. If a ticker already exists, the new entry REPLACES it.
function addPositions(newPositions) {
  for (const np of newPositions) {
    const idx = portfolio.findIndex(p => p.ticker === np.ticker);
    if (idx >= 0) portfolio[idx] = np;
    else portfolio.push(np);
  }
  saveMemoryToCloud();
  renderPortfolioPanel();
}

function removePosition(ticker) {
  portfolio = portfolio.filter(p => p.ticker !== ticker);
  saveMemoryToCloud();
  renderPortfolioPanel();
}

// ── INVESTOR PROFILE & ADVISOR ──────────────────────────────────────────────
// Profile-based screening rules — different filters per risk tolerance
// minDivYield is in PERCENT (e.g. 1.5 means 1.5%). The stored stk.dividendYield is in
// DECIMAL form (e.g. 0.023 means 2.3%) so we compare `dy * 100` against minDivYield.
const PROFILE_RULES = {
  conservative: {
    minHealth: 7.5, minGrowth: 0, minOverall: 7,
    maxPE: 28, minMarketCap: 50e9, maxDE: 1.2,
    requireDividend: true, minDivYield: 1.5,
    positionCap: 0.05, sectorCap: 0.25,
    description: 'Capital preservation focus: large-cap, manageable debt, dividend-paying quality names.'
  },
  moderate: {
    minHealth: 7, minGrowth: 6, minOverall: 7,
    maxPE: 40, minMarketCap: 20e9, maxDE: 1.5,
    requireDividend: false, minDivYield: 0,
    positionCap: 0.10, sectorCap: 0.35,
    description: 'Balance growth and quality: mid-to-large cap with reasonable valuations and diversified sectors.'
  },
  aggressive: {
    minHealth: 6, minGrowth: 7, minOverall: 6.5,
    maxPE: null, minMarketCap: 2e9, maxDE: null,
    requireDividend: false, minDivYield: 0,
    positionCap: 0.20, sectorCap: 0.50,
    description: 'Growth-focused, accepting higher volatility: smaller caps and richer valuations OK if growth is strong.'
  }
};

async function tool_recommendForProfile({ risk, exclude_sectors, exclude_owned }) {
  const rules = PROFILE_RULES[risk] || PROFILE_RULES.moderate;
  const excluded = new Set((exclude_sectors || []).map(s => s.toLowerCase()));
  const owned = (exclude_owned !== false && portfolio?.length) ? new Set(portfolio.map(p => p.ticker)) : new Set();

  const candidates = [];
  for (const t of TICKERS) {
    if (owned.has(t)) continue;
    const stk = STOCK[t];
    if (!stk || !ANNUAL[t]?.length) continue;
    if (excluded.has((stk.sector || '').toLowerCase())) continue;

    const m = calcMetrics(ANNUAL[t]);
    const scores = calcScores(m);

    if (scores.overall < rules.minOverall) continue;
    if (scores.health < rules.minHealth) continue;
    if (scores.growth < rules.minGrowth) continue;
    if (rules.maxPE != null && stk.pe > 0 && stk.pe > rules.maxPE) continue;
    if (rules.minMarketCap > 0 && (!stk.marketCap || stk.marketCap < rules.minMarketCap)) continue;
    const dte = m.lev?.slice(-1)[0]?.dte;
    if (rules.maxDE != null && dte != null && dte > rules.maxDE) continue;
    const dyPct = (stk.dividendYield || 0) * 100;
    if (rules.requireDividend && dyPct < rules.minDivYield) continue;

    candidates.push({
      ticker: t,
      sector: stk.sector,
      market_cap: stk.marketCap,
      price: stk.price ? parseFloat(stk.price.toFixed(2)) : null,
      pe: stk.pe > 0 ? parseFloat(stk.pe.toFixed(1)) : null,
      dividend_yield_pct: dyPct ? parseFloat(dyPct.toFixed(2)) : null,
      overall_score: scores.overall,
      growth_score: scores.growth,
      profit_score: scores.profitability,
      health_score: scores.health,
      cashflow_score: scores.cashflow,
      latest_revenue_growth_pct: m.yoy?.slice(-1)[0]?.revenue_growth ? parseFloat(m.yoy.slice(-1)[0].revenue_growth.toFixed(1)) : null,
      pct_from_52w_high: (stk.price && stk.fiftyTwoWeekHigh) ? parseFloat(((stk.price - stk.fiftyTwoWeekHigh) / stk.fiftyTwoWeekHigh * 100).toFixed(1)) : null
    });
  }

  candidates.sort((a, b) => b.overall_score - a.overall_score);

  // Enrich the top 8 with live news / earnings / analyst / KG context, in parallel
  const TOP_N_ENRICH = 8;
  const toEnrich = candidates.slice(0, TOP_N_ENRICH);
  const earningsMapPromise = fetchUpcomingEarningsMap();
  const enrichedTop = await Promise.all(toEnrich.map(c => enrichRecommendationCandidate(c)));
  const earningsMap = await earningsMapPromise;
  enrichedTop.forEach(c => {
    if (earningsMap[c.ticker]) c.next_earnings = earningsMap[c.ticker];
  });

  // Compute portfolio's thematic exposure from KG
  const portfolioThemes = await computePortfolioThematicExposure();

  // Group top by sector for diversification suggestions
  const bySector = {};
  candidates.slice(0, 30).forEach(c => { (bySector[c.sector] = bySector[c.sector] || []).push(c.ticker); });

  return {
    profile: risk,
    rules_applied: rules,
    candidate_count: candidates.length,
    top_candidates: enrichedTop,
    other_candidates: candidates.slice(TOP_N_ENRICH, 20),
    candidates_by_sector: bySector,
    portfolio_thematic_exposure: portfolioThemes,
    enrichment_notes: {
      top_n_enriched: enrichedTop.length,
      enrichments_included: ['recent_news_headlines', 'next_earnings_date', 'analyst_consensus', 'kg_related_to_portfolio']
    }
  };
}

// Per-candidate enrichment: news + analyst ratings + KG portfolio-exposure check
async function enrichRecommendationCandidate(candidate) {
  const t = candidate.ticker;
  const [news, analyst, kgRelated] = await Promise.all([
    fetchAdvisorNews(t),
    fetchAdvisorAnalyst(t),
    fetchAdvisorKG(t)
  ]);

  // Check: do any of this candidate's KG-related tickers appear in the user's portfolio?
  // If yes, recommending this ticker DEEPENS existing thematic exposure (likely undesirable for diversification)
  const portfolioTickers = new Set((portfolio || []).map(p => p.ticker));
  const kgOverlapWithPortfolio = (kgRelated || [])
    .filter(r => portfolioTickers.has(r.related_ticker))
    .map(r => ({ ticker: r.related_ticker, type: r.type }));

  return {
    ...candidate,
    recent_news: news,
    analyst_consensus: analyst,
    kg_overlap_with_portfolio: kgOverlapWithPortfolio,
    thematic_caution: kgOverlapWithPortfolio.length > 0
      ? `WARNING: ${t} is connected to ${kgOverlapWithPortfolio.length} of your existing holdings (${kgOverlapWithPortfolio.map(x => `${x.ticker}:${x.type}`).join(', ')}). Adding ${t} may deepen existing thematic exposure rather than diversify.`
      : null
  };
}

async function fetchAdvisorNews(ticker) {
  if (!FMP_QUOTE_KEY) return [];
  try {
    const res = await fetch(`https://financialmodelingprep.com/stable/news/stock?symbols=${ticker}&limit=3&apikey=${FMP_QUOTE_KEY}`);
    if (!res.ok) return [];
    const data = await res.json();
    if (!Array.isArray(data)) return [];
    return data.slice(0, 3).map(n => ({
      title: n.title,
      date: n.publishedDate?.slice(0, 10),
      site: n.site || n.publisher,
      excerpt: (n.text || '').slice(0, 200)
    }));
  } catch { return []; }
}

async function fetchAdvisorAnalyst(ticker) {
  if (!FMP_QUOTE_KEY) return null;
  try {
    const [ratingRes, targetRes] = await Promise.all([
      fetch(`https://financialmodelingprep.com/stable/ratings-snapshot?symbol=${ticker}&apikey=${FMP_QUOTE_KEY}`),
      fetch(`https://financialmodelingprep.com/stable/price-target-consensus?symbol=${ticker}&apikey=${FMP_QUOTE_KEY}`)
    ]);
    let rating = null, target = null;
    if (ratingRes.ok) {
      const rd = await ratingRes.json();
      const r = Array.isArray(rd) ? rd[0] : rd;
      if (r) rating = {
        rating: r.rating || r.ratingRecommendation,
        score: r.ratingScore,
        details: r.ratingDetailsDCFRecommendation || null
      };
    }
    if (targetRes.ok) {
      const td = await targetRes.json();
      const t2 = Array.isArray(td) ? td[0] : td;
      if (t2) target = {
        avg: t2.targetConsensus ?? t2.targetMedian ?? t2.targetMean,
        high: t2.targetHigh,
        low: t2.targetLow,
        analyst_count: t2.targetMedianAnalysts ?? t2.numberOfAnalysts ?? null
      };
    }
    if (!rating && !target) return null;
    const result = {};
    if (rating) result.rating = rating;
    if (target) {
      result.price_target = target;
      const curPrice = STOCK[ticker]?.price;
      if (curPrice && target.avg) {
        result.upside_pct = parseFloat(((target.avg - curPrice) / curPrice * 100).toFixed(1));
      }
    }
    return result;
  } catch { return null; }
}

async function fetchAdvisorKG(ticker) {
  if (!supabaseClient) return [];
  try {
    const { data } = await supabaseClient
      .from('ticker_relationships')
      .select('related_ticker, type, evidence')
      .eq('source_ticker', ticker)
      .order('strength', { ascending: false })
      .limit(15);
    return data || [];
  } catch { return []; }
}

// One batch FMP call for upcoming earnings; cached for 1 hour
let _earningsMapCache = { data: null, fetchedAt: 0 };
async function fetchUpcomingEarningsMap() {
  if (!FMP_QUOTE_KEY) return {};
  const HOUR = 60 * 60 * 1000;
  if (_earningsMapCache.data && Date.now() - _earningsMapCache.fetchedAt < HOUR) return _earningsMapCache.data;
  try {
    const now = new Date();
    const inSixtyDays = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);
    const from = now.toISOString().slice(0, 10);
    const to = inSixtyDays.toISOString().slice(0, 10);
    const res = await fetch(`https://financialmodelingprep.com/stable/earnings-calendar?from=${from}&to=${to}&apikey=${FMP_QUOTE_KEY}`);
    if (!res.ok) return {};
    const data = await res.json();
    if (!Array.isArray(data)) return {};
    const map = {};
    for (const e of data) {
      if (!e.symbol || !e.date) continue;
      // Only keep the next (earliest) upcoming date per ticker
      if (!map[e.symbol] || e.date < map[e.symbol].date) {
        const daysUntil = Math.ceil((new Date(e.date) - now) / (24 * 60 * 60 * 1000));
        map[e.symbol] = { date: e.date, days_until: daysUntil, eps_estimate: e.epsEstimated ?? null, revenue_estimate: e.revenueEstimated ?? null };
      }
    }
    _earningsMapCache = { data: map, fetchedAt: Date.now() };
    return map;
  } catch { return {}; }
}

// What themes/relationships does the user's portfolio cluster around?
async function computePortfolioThematicExposure() {
  if (!portfolio?.length || !supabaseClient) return null;
  const tickers = portfolio.map(p => p.ticker);
  try {
    const { data } = await supabaseClient
      .from('ticker_relationships')
      .select('source_ticker, related_ticker, type')
      .in('source_ticker', tickers);
    if (!data?.length) return { note: 'No KG relationships seeded for held tickers.', themes: [] };
    // Count how often each (related_ticker, type) appears across held tickers
    const counts = {};
    for (const r of data) {
      const k = `${r.related_ticker}|${r.type}`;
      counts[k] = (counts[k] || 0) + 1;
    }
    // Themes the portfolio is concentrated in
    const themes = Object.entries(counts)
      .filter(([_, c]) => c >= 2)
      .sort(([_, a], [__, b]) => b - a)
      .slice(0, 10)
      .map(([k, count]) => {
        const [related, type] = k.split('|');
        return { related_ticker: related, relationship_type: type, count, interpretation: `${count} of your held tickers have a ${type} relationship with ${related}` };
      });
    return { themes_with_2plus_overlap: themes, total_kg_relationships: data.length };
  } catch { return null; }
}

// Agent tool: structured snapshot of the user's portfolio for the council
function tool_getPortfolio() {
  if (!portfolio?.length) {
    return { count: 0, message: 'Portfolio is empty. The user has not added any positions yet.' };
  }
  const data = calcPortfolio();
  if (!data) return { count: 0, message: 'Portfolio data unavailable' };
  return {
    count: data.count,
    total_value: parseFloat(data.totalValue.toFixed(2)),
    day_pl: data.totalDayPL != null ? parseFloat(data.totalDayPL.toFixed(2)) : null,
    day_pl_pct: data.totalDayPct != null ? parseFloat(data.totalDayPct.toFixed(2)) : null,
    total_unrealized_pl: data.totalUnrealized != null ? parseFloat(data.totalUnrealized.toFixed(2)) : null,
    total_unrealized_pct: data.totalUnrealizedPct != null ? parseFloat(data.totalUnrealizedPct.toFixed(2)) : null,
    weighted_quality_score: data.weightedScore != null ? parseFloat(data.weightedScore.toFixed(2)) : null,
    sector_allocation: data.sectorAlloc.map(s => ({
      sector: s.sector,
      pct: parseFloat(s.pct.toFixed(1)),
      value: parseFloat(s.value.toFixed(2))
    })),
    positions: data.positions.slice(0, 25).map(p => ({
      ticker: p.ticker,
      sector: p.sector,
      shares: p.shares,
      cost_basis: p.costBasis,
      current_price: p.price ? parseFloat(p.price.toFixed(2)) : null,
      value: p.value ? parseFloat(p.value.toFixed(2)) : null,
      unrealized_pct: p.unrealizedPct != null ? parseFloat(p.unrealizedPct.toFixed(2)) : null,
      day_pct: p.dayPct != null ? parseFloat(p.dayPct.toFixed(2)) : null,
      quality_score: p.score
    }))
  };
}

// Refresh live quotes for all portfolio tickers. Throttled to once per 30s.
const PORTFOLIO_QUOTE_TTL_MS = 30 * 1000;
let _portfolioQuotesRefreshing = false;
async function refreshPortfolioQuotes() {
  if (_portfolioQuotesRefreshing || !portfolio.length || !FMP_QUOTE_KEY) return;
  // Only refresh tickers whose price is stale (>30s old)
  const now = Date.now();
  const stale = portfolio
    .map(p => p.ticker)
    .filter(t => {
      const stk = STOCK[t];
      return !stk?._priceUpdatedAt || (now - stk._priceUpdatedAt) > PORTFOLIO_QUOTE_TTL_MS;
    });
  if (!stale.length) return;
  _portfolioQuotesRefreshing = true;
  try {
    // FMP /stable/quote is single-symbol only. Fan out in parallel, max 5 concurrent.
    const BATCH = 5;
    for (let i = 0; i < stale.length; i += BATCH) {
      const slice = stale.slice(i, i + BATCH);
      const quotes = await Promise.all(slice.map(t => fetchLiveQuote(t).catch(() => null)));
      slice.forEach((t, idx) => {
        const q = quotes[idx];
        if (q) updateStockBarWithLive(t, q); // already persists price/changePct/pe/52w + sets _priceUpdatedAt
      });
      if (i + BATCH < stale.length) await new Promise(r => setTimeout(r, 150));
    }
  } catch (e) {
    console.warn('Portfolio quote refresh failed:', e);
  } finally {
    _portfolioQuotesRefreshing = false;
  }
}

// Sidebar summary render. Renders immediately with cached data, then refreshes live quotes
// in the background and re-renders with fresh prices when they arrive.
function renderPortfolioPanel() {
  renderPortfolioPanelInner();
  refreshPortfolioQuotes()
    .then(() => renderPortfolioPanelInner())
    .catch(() => {});
}

function renderPortfolioPanelInner() {
  const el = document.getElementById('portfolioPanel');
  if (!el) return;
  const isAr = lang === 'ar';
  const data = calcPortfolio();
  if (!data) {
    el.innerHTML = `<div class="port-empty">
      <div style="font-size:11px;color:var(--text3);padding:6px 4px">
        ${isAr ? 'لا توجد مراكز بعد' : 'No positions yet'}
      </div>
      <button class="port-add-btn" onclick="openPortfolioModal()">+ ${isAr ? 'إضافة مراكز' : 'Add positions'}</button>
    </div>`;
    return;
  }
  const dayClass = data.totalDayPL > 0 ? 'pos' : data.totalDayPL < 0 ? 'neg' : 'neu';
  const totalClass = data.totalUnrealized > 0 ? 'pos' : data.totalUnrealized < 0 ? 'neg' : 'neu';
  el.innerHTML = `<div class="port-summary" onclick="showPortfolioDashboard()">
    <div class="port-row"><span class="port-lbl">${isAr ? 'القيمة' : 'Value'}</span><span class="port-val">${fM(data.totalValue)}</span></div>
    ${data.totalDayPL != null ? `<div class="port-row"><span class="port-lbl">${isAr ? 'اليوم' : 'Day'}</span><span class="port-val ${dayClass}">${data.totalDayPL > 0 ? '+' : ''}${fM(Math.abs(data.totalDayPL))} (${data.totalDayPct > 0 ? '+' : ''}${data.totalDayPct.toFixed(2)}%)</span></div>` : ''}
    ${data.totalUnrealized != null ? `<div class="port-row"><span class="port-lbl">${isAr ? 'إجمالي' : 'Total'}</span><span class="port-val ${totalClass}">${data.totalUnrealized > 0 ? '+' : ''}${fM(Math.abs(data.totalUnrealized))} (${data.totalUnrealizedPct > 0 ? '+' : ''}${data.totalUnrealizedPct.toFixed(1)}%)</span></div>` : ''}
    <div class="port-row"><span class="port-lbl">${data.count} ${isAr ? 'مراكز' : 'positions'}</span>${data.weightedScore != null ? `<span class="port-score" style="color:${sCol(data.weightedScore)}">★ ${data.weightedScore.toFixed(1)}</span>` : ''}</div>
  </div>
  <button class="port-add-btn-mini" onclick="openPortfolioModal()" title="${isAr ? 'إضافة' : 'Add positions'}">+ ${isAr ? 'إضافة' : 'Add'}</button>`;
}

// Interactive modal — ticker autocomplete + per-position add/remove (no CSV burden)
function openPortfolioModal() {
  const isAr = lang === 'ar';
  const modal = document.createElement('div');
  modal.className = 'port-modal-bg';
  modal.innerHTML = `<div class="port-modal">
    <div class="port-modal-hdr">
      <h3>💼 ${isAr ? 'إدارة المحفظة' : 'Manage Portfolio'}</h3>
      <button class="port-modal-close" onclick="this.closest('.port-modal-bg').remove()">✕</button>
    </div>
    <div class="port-modal-body">

      <div class="port-add-form">
        <div class="port-add-label">${isAr ? 'إضافة مركز جديد' : 'Add a position'}</div>
        <div class="port-search-wrap">
          <input id="portTickerInput" type="text" class="port-input" placeholder="${isAr ? '🔍 ابحث بالرمز أو الاسم...' : '🔍 Search ticker or company...'}" autocomplete="off" />
          <div id="portTickerDrop" class="port-search-drop"></div>
        </div>
        <div id="portSelectedInfo" class="port-selected hidden">
          <span class="port-selected-ticker"></span>
          <span class="port-selected-name"></span>
          <span class="port-selected-price"></span>
        </div>
        <div class="port-row-inputs">
          <input id="portSharesInput" type="number" class="port-input port-input-small" placeholder="${isAr ? 'عدد الأسهم' : 'Shares'}" min="0" step="any" />
          <input id="portCostInput" type="number" class="port-input port-input-small" placeholder="${isAr ? 'متوسط التكلفة (اختياري)' : 'Cost basis (optional)'}" min="0" step="any" />
          <button id="portAddBtn" class="port-save-btn port-add-position-btn" onclick="addPositionFromForm()" disabled>+ ${isAr ? 'إضافة' : 'Add'}</button>
        </div>
        <div id="portFormStatus" class="port-status"></div>
      </div>

      <div class="port-divider"></div>

      <div class="port-list-section">
        <div class="port-add-label">${isAr ? 'مراكزك الحالية' : 'Your positions'} (<span id="portCount">${portfolio.length}</span>)</div>
        <div id="portList" class="port-list"></div>
      </div>

      <div class="port-modal-actions">
        <button class="port-save-btn" onclick="this.closest('.port-modal-bg').remove()">${isAr ? 'تم' : 'Done'}</button>
      </div>
    </div>
  </div>`;
  document.body.appendChild(modal);
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });

  // Refresh live quotes for all positions, then re-render the list with fresh prices
  refreshPortfolioQuotes().then(() => renderPortfolioList()).catch(() => {});

  // Wire up the ticker search input
  const input = document.getElementById('portTickerInput');
  input.addEventListener('input', () => portTickerSearch(input.value));
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { input.value = ''; portTickerSearch(''); }
    if (e.key === 'Enter') {
      const first = document.querySelector('#portTickerDrop .port-drop-item');
      if (first) first.click();
    }
  });
  input.focus();

  renderPortfolioList();
}

// Stores the currently-selected ticker in the form before "Add" is clicked
let _portSelectedTicker = null;

function portTickerSearch(q) {
  const drop = document.getElementById('portTickerDrop');
  if (!drop) return;
  const query = (q || '').trim().toUpperCase();
  if (!query) { drop.innerHTML = ''; drop.classList.remove('open'); return; }
  // Score each match: exact ticker (best) > ticker prefix > company name. Then by market cap.
  const scored = [];
  for (const t of TICKERS) {
    const name = (STOCK[t]?.companyName || '').toUpperCase();
    let score = 0;
    if (t === query) score = 100;
    else if (t.startsWith(query)) score = 50;
    else if (name.startsWith(query)) score = 30;
    else if (name.includes(query)) score = 10;
    if (score > 0) scored.push({ t, score, mcap: STOCK[t]?.marketCap || 0 });
  }
  scored.sort((a, b) => b.score - a.score || b.mcap - a.mcap);
  const matches = scored.slice(0, 10).map(x => x.t);
  if (!matches.length) {
    drop.innerHTML = `<div class="port-drop-empty">No matches</div>`;
    drop.classList.add('open');
    return;
  }
  drop.innerHTML = matches.map(t => {
    const stk = STOCK[t] || {};
    return `<div class="port-drop-item" onclick="portSelectTicker('${t}')">
      <span class="port-drop-ticker">${t}</span>
      <span class="port-drop-name">${stk.companyName || ''}</span>
      <span class="port-drop-price">${stk.price ? '$' + stk.price.toFixed(2) : ''}</span>
    </div>`;
  }).join('');
  drop.classList.add('open');
}

function portSelectTicker(ticker) {
  _portSelectedTicker = ticker;
  const stk = STOCK[ticker] || {};
  const input = document.getElementById('portTickerInput');
  if (input) input.value = ticker;
  document.getElementById('portTickerDrop')?.classList.remove('open');
  const info = document.getElementById('portSelectedInfo');
  const renderPill = () => {
    if (!info) return;
    const s = STOCK[ticker] || {};
    info.classList.remove('hidden');
    info.querySelector('.port-selected-ticker').textContent = ticker;
    info.querySelector('.port-selected-name').textContent = s.companyName || '';
    info.querySelector('.port-selected-price').textContent = s.price ? `$${s.price.toFixed(2)}` : '';
  };
  renderPill();
  // If the price is stale (>30s) or missing, fetch a fresh live quote and re-render the pill
  const stale = !stk._priceUpdatedAt || (Date.now() - stk._priceUpdatedAt) > PORTFOLIO_QUOTE_TTL_MS;
  if (stale && FMP_QUOTE_KEY) {
    fetchLiveQuote(ticker).then(q => {
      if (!q) return;
      updateStockBarWithLive(ticker, q);
      renderPill();
    }).catch(() => {});
  }
  // Pre-fill from existing position if it exists
  const existing = portfolio.find(p => p.ticker === ticker);
  const sharesInput = document.getElementById('portSharesInput');
  const costInput = document.getElementById('portCostInput');
  if (existing) {
    if (sharesInput) sharesInput.value = existing.shares;
    if (costInput) costInput.value = existing.costBasis || '';
  } else {
    if (sharesInput) sharesInput.value = '';
    if (costInput) costInput.value = '';
  }
  document.getElementById('portAddBtn').disabled = false;
  sharesInput?.focus();
}

function addPositionFromForm() {
  const isAr = lang === 'ar';
  const status = document.getElementById('portFormStatus');
  if (!_portSelectedTicker) {
    if (status) status.innerHTML = `<span class="port-err">${isAr ? 'اختر رمزاً أولاً' : 'Pick a ticker first'}</span>`;
    return;
  }
  const shares = parseFloat(document.getElementById('portSharesInput').value);
  if (!isFinite(shares) || shares <= 0) {
    if (status) status.innerHTML = `<span class="port-err">${isAr ? 'أدخل عدد أسهم صحيح' : 'Enter valid shares'}</span>`;
    return;
  }
  const costRaw = document.getElementById('portCostInput').value;
  const costBasis = costRaw && isFinite(parseFloat(costRaw)) && parseFloat(costRaw) > 0 ? parseFloat(costRaw) : null;

  addPositions([{
    ticker: _portSelectedTicker,
    shares,
    costBasis,
    addedAt: new Date().toISOString()
  }]);

  // Reset the form
  if (status) status.innerHTML = `<span class="port-ok">✓ ${_portSelectedTicker} ${isAr ? 'أضيف' : 'added'}</span>`;
  setTimeout(() => { if (status) status.innerHTML = ''; }, 2000);
  _portSelectedTicker = null;
  document.getElementById('portTickerInput').value = '';
  document.getElementById('portSharesInput').value = '';
  document.getElementById('portCostInput').value = '';
  document.getElementById('portSelectedInfo')?.classList.add('hidden');
  document.getElementById('portAddBtn').disabled = true;
  document.getElementById('portTickerInput').focus();
  renderPortfolioList();
}

function renderPortfolioList() {
  const el = document.getElementById('portList');
  const countEl = document.getElementById('portCount');
  if (!el) return;
  if (countEl) countEl.textContent = portfolio.length;
  if (!portfolio.length) {
    el.innerHTML = `<div class="port-list-empty">${lang==='ar' ? 'لا توجد مراكز بعد. أضف رمزاً أعلاه.' : 'No positions yet. Add one above.'}</div>`;
    return;
  }
  const isAr = lang === 'ar';
  // Sort by ticker A→Z for stability
  const sorted = [...portfolio].sort((a, b) => a.ticker.localeCompare(b.ticker));
  el.innerHTML = sorted.map(p => {
    const stk = STOCK[p.ticker] || {};
    const value = stk.price ? stk.price * p.shares : null;
    const cost = p.costBasis ? p.costBasis * p.shares : null;
    const pl = cost != null && value != null ? value - cost : null;
    const plPct = cost && value ? ((value - cost) / cost) * 100 : null;
    const plClass = pl == null ? 'neu' : pl > 0 ? 'pos' : pl < 0 ? 'neg' : 'neu';
    return `<div class="port-list-item">
      <div class="port-list-main" onclick="portSelectTicker('${p.ticker}')">
        <span class="port-list-ticker">${p.ticker}</span>
        <span class="port-list-shares">${p.shares} ${isAr ? 'سهم' : 'shares'}${p.costBasis ? ` @ $${p.costBasis.toFixed(2)}` : ''}</span>
      </div>
      <div class="port-list-side">
        ${value != null ? `<span class="port-list-value">${fM(value)}</span>` : ''}
        ${pl != null ? `<span class="port-list-pl ${plClass}">${pl > 0 ? '+' : ''}${plPct.toFixed(1)}%</span>` : ''}
        <button class="port-list-del" title="${isAr ? 'حذف' : 'Remove'}" onclick="event.stopPropagation(); removePositionAndRefresh('${p.ticker}')">✕</button>
      </div>
    </div>`;
  }).join('');
}

function removePositionAndRefresh(ticker) {
  removePosition(ticker);
  renderPortfolioList();
}

// Full portfolio dashboard — modal with summary cards + holdings table + refresh
function showPortfolioDashboard() {
  if (!portfolio.length) return openPortfolioModal();
  const isAr = lang === 'ar';
  const modal = document.createElement('div');
  modal.className = 'port-modal-bg port-dash-bg';
  modal.innerHTML = `<div class="port-dash">
    <div class="port-dash-hdr">
      <div class="port-dash-title">💼 ${isAr ? 'محفظتي' : 'My Portfolio'}</div>
      <div class="port-dash-actions">
        <button class="port-dash-btn" id="portRefreshBtn" onclick="refreshPortfolioDash()" title="${isAr?'تحديث الأسعار':'Refresh prices'}">⟳ ${isAr ? 'تحديث' : 'Refresh'}</button>
        <button class="port-dash-btn" onclick="this.closest('.port-modal-bg').remove(); openPortfolioModal();">✏️ ${isAr ? 'تعديل' : 'Edit'}</button>
        <button class="port-dash-btn port-dash-ai" onclick="askCouncilAboutPortfolio(); this.closest('.port-modal-bg').remove();">🤖 ${isAr ? 'اسأل المجلس' : 'Ask Council'}</button>
        <button class="port-dash-btn port-dash-advisor" onclick="this.closest('.port-modal-bg').remove(); startPortfolioAdvisor();">💡 ${isAr ? 'توصيات' : 'Recommendations'}</button>
        <button class="port-modal-close" onclick="this.closest('.port-modal-bg').remove()">✕</button>
      </div>
    </div>
    <div id="portDashBody" class="port-dash-body"></div>
  </div>`;
  document.body.appendChild(modal);
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });
  renderPortfolioDashBody();
  // Background refresh — fresh quotes for accurate display
  refreshPortfolioQuotes().then(() => renderPortfolioDashBody()).catch(() => {});
}

// Manual refresh button — force refetch even if cache is "fresh"
async function refreshPortfolioDash() {
  const btn = document.getElementById('portRefreshBtn');
  if (btn) { btn.disabled = true; btn.textContent = '⟳ ' + (lang==='ar' ? 'يحدّث...' : 'Refreshing...'); }
  // Invalidate cache for ALL portfolio tickers, then refresh
  portfolio.forEach(p => { if (STOCK[p.ticker]) STOCK[p.ticker]._priceUpdatedAt = 0; });
  await refreshPortfolioQuotes();
  renderPortfolioDashBody();
  renderPortfolioPanelInner();
  if (btn) { btn.disabled = false; btn.textContent = '⟳ ' + (lang==='ar' ? 'تحديث' : 'Refresh'); }
}

// Sort state lives at module scope so it persists across re-renders within a dashboard session
let _portDashSort = { key: 'value', dir: 'desc' };

function setPortDashSort(key) {
  if (_portDashSort.key === key) {
    _portDashSort.dir = _portDashSort.dir === 'desc' ? 'asc' : 'desc';
  } else {
    _portDashSort = { key, dir: 'desc' };
  }
  renderPortfolioDashBody();
}

function renderPortfolioDashBody() {
  const el = document.getElementById('portDashBody');
  if (!el) return;
  const isAr = lang === 'ar';
  const data = calcPortfolio();
  if (!data) {
    el.innerHTML = `<div class="port-dash-empty">No positions</div>`;
    return;
  }

  // Sort positions by the active sort key
  const sortKey = _portDashSort.key;
  const sortDir = _portDashSort.dir === 'asc' ? 1 : -1;
  const sortVal = (p) => {
    switch (sortKey) {
      case 'ticker': return p.ticker;
      case 'value':  return p.value || 0;
      case 'price':  return p.price || 0;
      case 'pl':     return p.unrealized != null ? p.unrealized : -Infinity;
      case 'plPct':  return p.unrealizedPct != null ? p.unrealizedPct : -Infinity;
      case 'day':    return p.dayPL != null ? p.dayPL : -Infinity;
      default:       return 0;
    }
  };
  data.positions.sort((a, b) => {
    const av = sortVal(a), bv = sortVal(b);
    if (typeof av === 'string') return sortDir * av.localeCompare(bv);
    return sortDir * (av - bv);
  });

  // Summary cards
  const dayClass = data.totalDayPL > 0 ? 'pos' : data.totalDayPL < 0 ? 'neg' : 'neu';
  const totalClass = data.totalUnrealized > 0 ? 'pos' : data.totalUnrealized < 0 ? 'neg' : 'neu';
  const fmtDollar = (n) => (n == null) ? '—' : (n >= 0 ? '+' : '−') + '$' + Math.abs(n).toLocaleString(undefined, {maximumFractionDigits:2});
  const fmtPct = (n) => (n == null) ? '' : (n >= 0 ? '+' : '') + n.toFixed(2) + '%';
  const cards = `<div class="port-dash-cards">
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'القيمة الإجمالية' : 'Total Value'}</div>
      <div class="port-card-val">${fM(data.totalValue)}</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'ربح/خسارة اليوم' : 'Day P&L'}</div>
      <div class="port-card-val ${dayClass}">${data.totalDayPL != null ? fmtDollar(data.totalDayPL) : '—'}</div>
      <div class="port-card-sub ${dayClass}">${fmtPct(data.totalDayPct)}</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'الربح/الخسارة الإجمالي' : 'Total P&L'}</div>
      <div class="port-card-val ${totalClass}">${data.totalUnrealized != null ? fmtDollar(data.totalUnrealized) : '—'}</div>
      <div class="port-card-sub ${totalClass}">${fmtPct(data.totalUnrealizedPct)}</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'متوسط الجودة' : 'Avg Quality'}</div>
      <div class="port-card-val" style="color:${data.weightedScore != null ? sCol(data.weightedScore) : 'var(--text3)'}">${data.weightedScore != null ? '★ ' + data.weightedScore.toFixed(1) : '—'}</div>
      <div class="port-card-sub">${data.count} ${isAr ? 'مراكز' : 'positions'}</div>
    </div>
  </div>`;

  // Sector allocation card with mini pie chart
  const sectorBlock = `<div class="port-sector-card">
    <div class="port-sector-hdr">${isAr ? 'توزيع القطاعات' : 'Sector Allocation'}</div>
    <div class="port-sector-body">
      <div class="port-sector-canvas-wrap"><canvas id="portSectorPie"></canvas></div>
      <div class="port-sector-legend">
        ${data.sectorAlloc.slice(0, 8).map((s, i) => `
          <div class="port-sector-row">
            <span class="port-sector-dot" style="background:${SECTOR_COLORS[i % SECTOR_COLORS.length]}"></span>
            <span class="port-sector-name">${s.sector}</span>
            <span class="port-sector-pct">${s.pct.toFixed(1)}%</span>
          </div>`).join('')}
      </div>
    </div>
  </div>`;

  // Sortable headers
  const arrow = (key) => _portDashSort.key === key ? (_portDashSort.dir === 'asc' ? ' ↑' : ' ↓') : '';
  const tableHdr = `<div class="port-thead">
    <div class="port-tcell port-tcell-ticker port-th" onclick="setPortDashSort('ticker')">${isAr ? 'السهم' : 'Stock'}${arrow('ticker')}</div>
    <div class="port-tcell port-th" onclick="setPortDashSort('value')">${isAr ? 'القيمة / الكمية' : 'Value / Qty'}${arrow('value')}</div>
    <div class="port-tcell port-th" onclick="setPortDashSort('price')">${isAr ? 'السعر / التكلفة' : 'Price / Cost'}${arrow('price')}</div>
    <div class="port-tcell port-tcell-pl port-th" onclick="setPortDashSort('pl')">${isAr ? 'إجمالي ربح/خسارة' : 'Total P&L'}${arrow('pl')}</div>
    <div class="port-tcell port-tcell-day port-th" onclick="setPortDashSort('day')">${isAr ? 'ربح اليوم' : 'Day'}${arrow('day')}</div>
    <div class="port-tcell port-tcell-actions"></div>
  </div>`;

  // Table rows with inline edit + delete actions
  const rows = data.positions.map(p => {
    const stk = STOCK[p.ticker] || {};
    const plClass = p.unrealized == null ? 'neu' : p.unrealized > 0 ? 'pos' : p.unrealized < 0 ? 'neg' : 'neu';
    const dayClass2 = p.dayPL == null ? 'neu' : p.dayPL > 0 ? 'pos' : p.dayPL < 0 ? 'neg' : 'neu';
    return `<div class="port-trow">
      <div class="port-tcell port-tcell-ticker" onclick="onTickerClick('${p.ticker}'); this.closest('.port-modal-bg').remove();">
        <div class="port-tt-name">${stk.companyName ? stk.companyName.slice(0, 22) : p.ticker}</div>
        <div class="port-tt-sym">${p.ticker}</div>
      </div>
      <div class="port-tcell port-tcell-value" onclick="onTickerClick('${p.ticker}'); this.closest('.port-modal-bg').remove();">
        <div class="port-tc-main">${p.value != null ? fM(p.value).replace('$','') : '—'}</div>
        <div class="port-tc-sub port-editable" onclick="event.stopPropagation(); portEditShares('${p.ticker}', this)" title="${isAr?'انقر لتعديل الكمية':'Click to edit shares'}">${p.shares} ${isAr ? 'سهم' : 'sh'}</div>
      </div>
      <div class="port-tcell port-tcell-price" onclick="onTickerClick('${p.ticker}'); this.closest('.port-modal-bg').remove();">
        <div class="port-tc-main">${p.price ? p.price.toFixed(2) : '—'}</div>
        <div class="port-tc-sub port-editable" onclick="event.stopPropagation(); portEditCost('${p.ticker}', this)" title="${isAr?'انقر لتعديل التكلفة':'Click to edit cost basis'}">${p.costBasis ? p.costBasis.toFixed(2) : (isAr?'—':'set cost')}</div>
      </div>
      <div class="port-tcell port-tcell-pl">
        <div class="port-tc-main ${plClass}">${p.unrealized != null ? fmtDollar(p.unrealized).replace('$','') : '—'}</div>
        <div class="port-tc-sub ${plClass}">${p.unrealizedPct != null ? fmtPct(p.unrealizedPct) : ''}</div>
      </div>
      <div class="port-tcell port-tcell-day">
        <div class="port-tc-main ${dayClass2}">${p.dayPL != null ? fmtDollar(p.dayPL).replace('$','') : '—'}</div>
        <div class="port-tc-sub ${dayClass2}">${p.dayPct != null ? fmtPct(p.dayPct) : ''}</div>
      </div>
      <div class="port-tcell port-tcell-actions">
        <button class="port-row-del" title="${isAr ? 'حذف' : 'Remove'}" onclick="event.stopPropagation(); if(confirm('${isAr?'حذف':'Remove'} ${p.ticker}?')) { removePosition('${p.ticker}'); renderPortfolioDashBody(); }">✕</button>
      </div>
    </div>`;
  }).join('');

  el.innerHTML = cards + sectorBlock + `<div class="port-table">${tableHdr}${rows}</div>`;

  // Render the pie chart after DOM update
  setTimeout(() => renderSectorPie(data.sectorAlloc), 20);
}

// Sector colors (consistent palette for pie + legend)
const SECTOR_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#a855f7', '#06b6d4', '#ec4899', '#84cc16', '#6366f1', '#f97316', '#14b8a6', '#8b5cf6'];

let _portSectorChart = null;
function renderSectorPie(sectorAlloc) {
  const canvas = document.getElementById('portSectorPie');
  if (!canvas || typeof Chart === 'undefined') return;
  if (_portSectorChart) { try { _portSectorChart.destroy(); } catch {} _portSectorChart = null; }
  const top = sectorAlloc.slice(0, 8);
  _portSectorChart = new Chart(canvas.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: top.map(s => s.sector),
      datasets: [{
        data: top.map(s => s.value),
        backgroundColor: top.map((_, i) => SECTOR_COLORS[i % SECTOR_COLORS.length]),
        borderColor: 'transparent',
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.label}: ${(ctx.parsed / ctx.dataset.data.reduce((s,v)=>s+v,0) * 100).toFixed(1)}% (${fM(ctx.parsed)})`
          }
        }
      },
      cutout: '62%'
    }
  });
}

// Inline edit — turn a cell into an input on click
function portEditShares(ticker, cellEl) {
  const pos = portfolio.find(p => p.ticker === ticker);
  if (!pos) return;
  const oldText = cellEl.textContent;
  cellEl.innerHTML = `<input type="number" class="port-inline-input" value="${pos.shares}" step="any" min="0" />`;
  const inp = cellEl.querySelector('input');
  inp.focus();
  inp.select();
  const commit = () => {
    const v = parseFloat(inp.value);
    if (isFinite(v) && v > 0) {
      pos.shares = v;
      saveMemoryToCloud();
      renderPortfolioDashBody();
      renderPortfolioPanelInner();
    } else cellEl.textContent = oldText;
  };
  const cancel = () => { cellEl.textContent = oldText; };
  inp.addEventListener('blur', commit);
  inp.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { inp.removeEventListener('blur', commit); commit(); }
    if (e.key === 'Escape') { inp.removeEventListener('blur', commit); cancel(); }
  });
}

function portEditCost(ticker, cellEl) {
  const pos = portfolio.find(p => p.ticker === ticker);
  if (!pos) return;
  const oldText = cellEl.textContent;
  cellEl.innerHTML = `<input type="number" class="port-inline-input" value="${pos.costBasis || ''}" placeholder="cost" step="any" min="0" />`;
  const inp = cellEl.querySelector('input');
  inp.focus();
  inp.select();
  const commit = () => {
    const raw = inp.value;
    if (raw === '' || raw == null) { pos.costBasis = null; }
    else {
      const v = parseFloat(raw);
      if (isFinite(v) && v > 0) pos.costBasis = v;
      else { cellEl.textContent = oldText; return; }
    }
    saveMemoryToCloud();
    renderPortfolioDashBody();
    renderPortfolioPanelInner();
  };
  const cancel = () => { cellEl.textContent = oldText; };
  inp.addEventListener('blur', commit);
  inp.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { inp.removeEventListener('blur', commit); commit(); }
    if (e.key === 'Escape') { inp.removeEventListener('blur', commit); cancel(); }
  });
}

// Advisor flow: capture profile (or use saved) → run council in recommendation mode
function startPortfolioAdvisor() {
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">🔒 ${lang==='ar'?'سجل الدخول لاستخدام المستشار':'Sign in to use the advisor'}</div>`);
    return;
  }
  // Always show the modal so the user can confirm/edit profile before each run
  openInvestorProfileModal((profile) => {
    investorProfile = { ...profile, updatedAt: new Date().toISOString() };
    saveMemoryToCloud();
    runRecommendationCouncil();
  });
}

function openInvestorProfileModal(onConfirm) {
  const isAr = lang === 'ar';
  const cur = investorProfile;
  const allSectors = [...new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean))].sort();
  const modal = document.createElement('div');
  modal.className = 'port-modal-bg';
  modal.innerHTML = `<div class="port-modal" style="max-width:520px">
    <div class="port-modal-hdr">
      <h3>💡 ${isAr ? 'الحصول على توصيات' : 'Get Recommendations'}</h3>
      <button class="port-modal-close" onclick="this.closest('.port-modal-bg').remove()">✕</button>
    </div>
    <div class="port-modal-body" style="gap:14px">
      <div>
        <div class="port-add-label">${isAr ? 'مستوى المخاطرة' : 'Risk tolerance'}</div>
        <div class="risk-options">
          <label class="risk-opt ${cur.riskTolerance === 'conservative' ? 'selected' : ''}"><input type="radio" name="risk" value="conservative" ${cur.riskTolerance === 'conservative' ? 'checked' : ''}><div class="risk-opt-title">🛡️ ${isAr ? 'متحفظ' : 'Conservative'}</div><div class="risk-opt-sub">${isAr ? 'الحفاظ على رأس المال، أرباح موزعة' : 'Capital preservation, dividends'}</div></label>
          <label class="risk-opt ${cur.riskTolerance === 'moderate' ? 'selected' : ''}"><input type="radio" name="risk" value="moderate" ${cur.riskTolerance === 'moderate' || !cur.riskTolerance ? 'checked' : ''}><div class="risk-opt-title">⚖️ ${isAr ? 'متوازن' : 'Moderate'}</div><div class="risk-opt-sub">${isAr ? 'نمو وجودة متوازنة' : 'Balanced growth + quality'}</div></label>
          <label class="risk-opt ${cur.riskTolerance === 'aggressive' ? 'selected' : ''}"><input type="radio" name="risk" value="aggressive" ${cur.riskTolerance === 'aggressive' ? 'checked' : ''}><div class="risk-opt-title">🚀 ${isAr ? 'مغامر' : 'Aggressive'}</div><div class="risk-opt-sub">${isAr ? 'نمو عالٍ، تقلبات أعلى' : 'High growth, higher volatility'}</div></label>
        </div>
      </div>

      <div>
        <div class="port-add-label">${isAr ? 'النقد المتاح للاستثمار ($)' : 'Cash to invest ($)'}</div>
        <input id="profCashInput" type="number" class="port-input" min="100" step="100" value="${cur.cashAvailable || 10000}" placeholder="10000" />
      </div>

      <div>
        <div class="port-add-label">${isAr ? 'استبعاد قطاعات (اختياري)' : 'Exclude sectors (optional)'}</div>
        <div class="sector-chips">
          ${allSectors.map(s => `<label class="sector-chip ${(cur.excludedSectors || []).includes(s) ? 'selected' : ''}">
            <input type="checkbox" name="excl" value="${s}" ${(cur.excludedSectors || []).includes(s) ? 'checked' : ''}>${s}
          </label>`).join('')}
        </div>
      </div>

      <label class="div-toggle"><input id="profReqDiv" type="checkbox" ${cur.requireDividend ? 'checked' : ''}> ${isAr ? 'يجب أن تدفع توزيعات أرباح' : 'Must pay dividends'}</label>

      <div class="port-modal-actions">
        <button class="port-cancel-btn" onclick="this.closest('.port-modal-bg').remove()">${isAr ? 'إلغاء' : 'Cancel'}</button>
        <button class="port-save-btn" id="profileSubmit">🤖 ${isAr ? 'احصل على توصيات' : 'Get Recommendations'}</button>
      </div>
    </div>
  </div>`;
  document.body.appendChild(modal);

  // Radio styling
  modal.querySelectorAll('.risk-opt').forEach(label => {
    label.addEventListener('click', () => {
      modal.querySelectorAll('.risk-opt').forEach(l => l.classList.remove('selected'));
      label.classList.add('selected');
    });
  });
  // Sector chip styling
  modal.querySelectorAll('.sector-chip').forEach(label => {
    const cb = label.querySelector('input');
    cb.addEventListener('change', () => { label.classList.toggle('selected', cb.checked); });
  });

  modal.querySelector('#profileSubmit').addEventListener('click', () => {
    const risk = modal.querySelector('input[name="risk"]:checked')?.value || 'moderate';
    const cash = parseFloat(modal.querySelector('#profCashInput').value) || 10000;
    const excluded = [...modal.querySelectorAll('input[name="excl"]:checked')].map(cb => cb.value);
    const requireDividend = modal.querySelector('#profReqDiv').checked;
    modal.remove();
    onConfirm({ riskTolerance: risk, cashAvailable: cash, excludedSectors: excluded, requireDividend });
  });
}

function runRecommendationCouncil() {
  const p = investorProfile;
  const isAr = lang === 'ar';
  const portfolioCtx = portfolio.length ? calcPortfolio() : null;
  const portSummary = portfolioCtx
    ? `Current portfolio: $${portfolioCtx.totalValue.toFixed(0)} across ${portfolioCtx.count} positions (${portfolioCtx.positions.slice(0, 10).map(x => `${x.ticker} ${(x.value / portfolioCtx.totalValue * 100).toFixed(0)}%`).join(', ')}). Sector exposure: ${portfolioCtx.sectorAlloc.slice(0, 5).map(s => `${s.sector} ${s.pct.toFixed(0)}%`).join(', ')}.`
    : 'No current positions — building from scratch.';

  const message = `[ADVISOR MODE] Recommend NEW stock positions for this investor.

PROFILE:
- Risk tolerance: ${p.riskTolerance}
- Cash to deploy: $${p.cashAvailable.toLocaleString()}
- Excluded sectors: ${p.excludedSectors?.length ? p.excludedSectors.join(', ') : 'none'}
- Dividend required: ${p.requireDividend ? 'yes' : 'no'}

${portSummary}

YOUR TASK:
- Call recommend_for_profile FIRST. Its top_candidates field is already enriched with: recent_news (headlines from this week), next_earnings (with days_until), analyst_consensus (rating + price target + upside %), kg_overlap_with_portfolio (warns if pick deepens existing exposure), and thematic_caution flag.
- Call get_portfolio to check current sector concentration.
- Each specialist proposes 2-3 NEW positions with rationale from your lens.
- MANDATORY checks for each pick — cite the live data:
  • Recent news: cite an actual headline from recent_news if available; if news is mixed/bearish, say so or skip the pick.
  • Earnings: if next_earnings.days_until <= 7, add an "⚠️ Earnings in N days — consider waiting" warning.
  • Analyst consensus: cite the consensus rating + price target + upside_pct.
  • Thematic exposure: if thematic_caution is set, prefer a DIFFERENT pick or explicitly acknowledge the concentration risk.
- For each pick give: $ allocation (within profile position cap), share count at current price, key catalyst (from recent_news or analyst data — NOT from generic knowledge), main risk, what to watch.
- Prefer picks that diversify (fill underweight sectors AND avoid kg_overlap_with_portfolio warnings).
- Max 5 total positions across all specialists. Keep 10-30% cash buffer.

Be specific. Cite real news headlines and real analyst targets — do not invent.`;

  if (typeof runAgent === 'function') runAgent(message);
  else if (typeof runCouncil === 'function') runCouncil(message);
}

// Chat shortcut — pipes portfolio into the council
function askCouncilAboutPortfolio() {
  if (!portfolio.length) return;
  const data = calcPortfolio();
  const positions = data.positions
    .map(p => `${p.ticker}: ${p.shares} sh${p.costBasis ? ` @ $${p.costBasis.toFixed(2)} cost` : ''}, value ${fM(p.value)}, ${p.unrealizedPct != null ? (p.unrealizedPct >= 0 ? '+' : '') + p.unrealizedPct.toFixed(1) + '%' : 'no cost basis'}`)
    .join('; ');
  const msg = `Analyze my portfolio (total value ${fM(data.totalValue)}, ${data.count} positions): ${positions}. What are the key concentration risks, quality of holdings, biggest red flags, and what should I watch?`;
  if (typeof runAgent === 'function') runAgent(msg);
  else if (typeof runCouncil === 'function') runCouncil(msg);
}

// ── MARKET PULSE — REMOVED ─────────────────────────────────────────────────

// (Market Pulse code removed)


// ── DECISION JOURNAL ────────────────────────────────────────────────────────
// Snapshot AI stock picks at the moment of recommendation, track them vs SPY over time.
// Answers: "do my picks actually outperform doing nothing?"

const SPY_TICKER = 'SPY';

// Save a pick to the journal. Auto-fills entry_price + SPY benchmark snapshot.
async function saveToJournal({ ticker, direction = 'long', rationale = '', source = 'manual', source_query = '', confidence = null }) {
  if (!isAiReady()) { addChatMessage?.('agent', '🔒 Sign in to save picks to your journal.'); return null; }
  if (!ticker) { console.error('saveToJournal: ticker required'); return null; }
  if (!supabaseClient) { console.error('Supabase not initialized'); return null; }

  const t = ticker.toUpperCase();

  // Fetch current price (live, not cached) + SPY for benchmark
  let entryPrice = null, spyPrice = null;
  try {
    const [q, spy] = await Promise.all([
      fetchLiveQuote(t),
      fetchLiveQuote(SPY_TICKER)
    ]);
    entryPrice = q?.price ?? STOCK[t]?.price ?? null;
    spyPrice = spy?.price ?? null;
  } catch (e) {
    console.warn('Live quote fetch failed, falling back to cached:', e);
    entryPrice = STOCK[t]?.price ?? null;
  }
  if (!entryPrice) { console.error('Could not determine entry price for', t); return null; }

  const row = {
    pin_hash: USER_AUTH_HASH,
    ticker: t,
    entry_price: entryPrice,
    spy_entry_price: spyPrice,
    direction,
    confidence,
    rationale: (rationale || '').slice(0, 2000),
    source,
    source_query: (source_query || '').slice(0, 500),
    closed: false
  };

  const { data, error } = await supabaseClient.from('decision_journal').insert(row).select().single();
  if (error) {
    console.error('Journal save failed:', error);
    return null;
  }
  console.log(`%c[Journal] ✓ Saved ${t} at $${entryPrice.toFixed(2)} (SPY $${spyPrice?.toFixed(2) ?? '—'})`, 'color:#10b981;font-weight:bold');
  updateJournalBadge();
  return data;
}

// Console helper: load all of user's journal entries
async function loadJournal() {
  if (!isAiReady() || !supabaseClient) return [];
  const { data, error } = await supabaseClient
    .from('decision_journal')
    .select('*')
    .eq('pin_hash', USER_AUTH_HASH)
    .order('saved_at', { ascending: false });
  if (error) { console.error(error); return []; }
  return data || [];
}

// Compute live performance for a single entry: ticker return, SPY return, alpha, days held
async function computeJournalRow(entry) {
  const t = entry.ticker;
  const currentPrice = STOCK[t]?.price ?? null; // already refreshed by computeAllJournalPerf
  const currentSpy = STOCK[SPY_TICKER]?.price ?? null;

  const tickerReturn = (currentPrice && entry.entry_price) ? ((currentPrice - entry.entry_price) / entry.entry_price * 100) : null;
  const spyReturn = (currentSpy && entry.spy_entry_price) ? ((currentSpy - entry.spy_entry_price) / entry.spy_entry_price * 100) : null;
  const alpha = (tickerReturn != null && spyReturn != null) ? (tickerReturn - spyReturn) : null;
  const daysHeld = Math.max(1, Math.floor((Date.now() - new Date(entry.saved_at).getTime()) / 86400000));

  return {
    ...entry,
    current_price: currentPrice,
    current_spy: currentSpy,
    ticker_return_pct: tickerReturn,
    spy_return_pct: spyReturn,
    alpha_pct: alpha,
    days_held: daysHeld,
    is_winner: alpha != null ? alpha > 0 : null
  };
}

// Fetch live SPY + all journal-ticker prices in parallel, then compute perf for each entry
async function computeAllJournalPerf(entries) {
  if (!entries?.length) return [];
  // Collect unique tickers (including SPY)
  const tickers = [...new Set([SPY_TICKER, ...entries.map(e => e.ticker)])];
  // Force a fresh-quote pass: only refresh ones not refreshed in last 60s
  const stale = tickers.filter(t => {
    const stk = STOCK[t];
    return !stk?._priceUpdatedAt || Date.now() - stk._priceUpdatedAt > 60000;
  });
  if (stale.length && FMP_QUOTE_KEY) {
    const BATCH = 5;
    for (let i = 0; i < stale.length; i += BATCH) {
      const slice = stale.slice(i, i + BATCH);
      const quotes = await Promise.all(slice.map(t => fetchLiveQuote(t).catch(() => null)));
      slice.forEach((t, idx) => { if (quotes[idx]) updateStockBarWithLive(t, quotes[idx]); });
      if (i + BATCH < stale.length) await new Promise(r => setTimeout(r, 120));
    }
  }
  return Promise.all(entries.map(e => computeJournalRow(e)));
}

// Mark an entry as closed at current price
async function closeJournalEntry(id) {
  if (!isAiReady() || !supabaseClient) return false;
  const entries = await loadJournal();
  const entry = entries.find(e => e.id === id);
  if (!entry) return false;
  const q = await fetchLiveQuote(entry.ticker);
  const exitPrice = q?.price ?? STOCK[entry.ticker]?.price ?? null;
  const { error } = await supabaseClient
    .from('decision_journal')
    .update({ closed: true, exit_price: exitPrice, exit_at: new Date().toISOString() })
    .eq('id', id)
    .eq('pin_hash', USER_AUTH_HASH);
  if (error) { console.error(error); return false; }
  return true;
}

async function deleteJournalEntry(id) {
  if (!isAiReady() || !supabaseClient) return false;
  const { error } = await supabaseClient
    .from('decision_journal')
    .delete()
    .eq('id', id)
    .eq('pin_hash', USER_AUTH_HASH);
  if (error) { console.error(error); return false; }
  updateJournalBadge();
  return true;
}

// Refresh the count badge in the sidebar "My Picks" label
async function updateJournalBadge() {
  const badge = document.getElementById('journalCountBadge');
  if (!badge || !isAiReady()) return;
  try {
    const { count } = await supabaseClient
      .from('decision_journal')
      .select('*', { count: 'exact', head: true })
      .eq('pin_hash', USER_AUTH_HASH);
    if (count != null && count > 0) {
      badge.textContent = count;
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  } catch { /* swallow */ }
}

// Wired to the "📓 Save" button on every ticker dashboard.
// Captures the currently-viewed ticker + most recent chat agent message as rationale.
async function quickSaveCurrentTicker() {
  if (!activeTicker) {
    console.warn('No active ticker to save');
    return;
  }
  const t = activeTicker;
  // Rationale = rule-based snapshot verdict (if computed for this ticker) + last AI take
  let rationale = '';
  if (_lastVerdict?.ticker === t) rationale = _lastVerdict.summary;
  const recentChat = (chatHistory || []).filter(m => m.role === 'agent' && m.ticker === t).pop();
  if (recentChat?.content) {
    const aiTxt = recentChat.content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 1200);
    rationale = rationale ? `${rationale}\n\nAI take: ${aiTxt}` : aiTxt;
  }
  const btn = document.querySelector('.sb-journal-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ ' + (lang==='ar'?'جاري الحفظ...':'Saving...'); }
  const result = await saveToJournal({ ticker: t, source: 'dashboard', rationale,
    confidence: _lastVerdict?.ticker === t ? _lastVerdict.confidence : null });
  if (btn) {
    if (result) {
      btn.textContent = '✓ ' + (lang==='ar'?'محفوظ':'Saved');
      btn.style.background = 'rgba(16,185,129,.15)';
      btn.style.borderColor = 'rgba(16,185,129,.4)';
      btn.style.color = '#10b981';
      setTimeout(() => {
        btn.textContent = '📓 ' + (lang==='ar'?'احفظ':'Save');
        btn.style.background = '';
        btn.style.borderColor = '';
        btn.style.color = '';
        btn.disabled = false;
      }, 2500);
    } else {
      btn.textContent = '❌ ' + (lang==='ar'?'فشل':'Failed');
      setTimeout(() => {
        btn.textContent = '📓 ' + (lang==='ar'?'احفظ':'Save');
        btn.disabled = false;
      }, 2500);
    }
  }
}

// Open the journal view modal
async function showJournal() {
  if (!isAiReady()) {
    if (typeof showPinModal === 'function') showPinModal();
    return;
  }
  const isAr = lang === 'ar';
  // Show modal with loading state first
  const modal = document.createElement('div');
  modal.className = 'port-modal-bg';
  modal.innerHTML = `<div class="port-dash">
    <div class="port-dash-hdr">
      <div class="port-dash-title">📓 ${isAr ? 'سجل توصياتي' : 'Decision Journal'}</div>
      <div class="port-dash-actions">
        <button class="port-dash-btn" onclick="refreshJournalView()" title="${isAr?'تحديث':'Refresh'}">⟳ ${isAr ? 'تحديث' : 'Refresh'}</button>
        <button class="port-modal-close" onclick="this.closest('.port-modal-bg').remove()">✕</button>
      </div>
    </div>
    <div id="journalBody" class="port-dash-body">
      <div class="port-dash-empty">⏳ ${isAr ? 'تحميل...' : 'Loading...'}</div>
    </div>
  </div>`;
  document.body.appendChild(modal);
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });
  await renderJournalBody();
}

async function refreshJournalView() { await renderJournalBody(); }

async function renderJournalBody() {
  const el = document.getElementById('journalBody');
  if (!el) return;
  const isAr = lang === 'ar';
  el.innerHTML = `<div class="port-dash-empty">⏳ ${isAr ? 'تحميل البيانات الحية...' : 'Loading live performance...'}</div>`;
  const entries = await loadJournal();
  if (!entries.length) {
    el.innerHTML = `<div class="port-dash-empty" style="padding:40px 20px">
      <div style="font-size:42px;margin-bottom:14px">📓</div>
      <div style="font-size:14px;color:var(--text);font-weight:600;margin-bottom:6px">${isAr ? 'السجل فارغ' : 'Your journal is empty'}</div>
      <div style="font-size:12px;color:var(--text3);max-width:380px;margin:0 auto;line-height:1.5">${isAr ? 'احفظ التوصيات من لوحة معلومات السهم لتتبع أدائها مقابل SPY بمرور الوقت.' : 'Save picks from any ticker dashboard to track their performance vs SPY over time. The aggregate stats will tell you whether your AI picks generate alpha.'}</div>
    </div>`;
    return;
  }
  const rows = await computeAllJournalPerf(entries);
  // Compute aggregate stats
  const closedRows = rows.filter(r => r.ticker_return_pct != null);
  const winners = closedRows.filter(r => r.is_winner);
  const losers = closedRows.filter(r => r.is_winner === false);
  const winRate = closedRows.length ? (winners.length / closedRows.length * 100) : 0;
  const avgAlpha = closedRows.length ? closedRows.reduce((s, r) => s + (r.alpha_pct || 0), 0) / closedRows.length : 0;
  const avgReturn = closedRows.length ? closedRows.reduce((s, r) => s + (r.ticker_return_pct || 0), 0) / closedRows.length : 0;
  const avgSpyReturn = closedRows.length ? closedRows.reduce((s, r) => s + (r.spy_return_pct || 0), 0) / closedRows.length : 0;
  const openCount = entries.filter(e => !e.closed).length;
  const alphaClass = avgAlpha > 0 ? 'pos' : avgAlpha < 0 ? 'neg' : 'neu';

  const fmtPct = (n, dp = 2) => n == null ? '—' : (n >= 0 ? '+' : '') + n.toFixed(dp) + '%';
  const cls = (n) => n == null ? 'neu' : n > 0 ? 'pos' : n < 0 ? 'neg' : 'neu';

  const cards = `<div class="port-dash-cards">
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'إجمالي التوصيات' : 'Total Picks'}</div>
      <div class="port-card-val">${entries.length}</div>
      <div class="port-card-sub">${openCount} ${isAr ? 'مفتوح' : 'open'} · ${entries.length - openCount} ${isAr ? 'مغلق' : 'closed'}</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'نسبة الفوز' : 'Win Rate'}</div>
      <div class="port-card-val ${winRate >= 50 ? 'pos' : 'neg'}">${winRate.toFixed(0)}%</div>
      <div class="port-card-sub">${winners.length}W / ${losers.length}L</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'متوسط العائد' : 'Avg Return'}</div>
      <div class="port-card-val ${cls(avgReturn)}">${fmtPct(avgReturn)}</div>
      <div class="port-card-sub">${isAr ? 'SPY:' : 'SPY:'} ${fmtPct(avgSpyReturn)}</div>
    </div>
    <div class="port-card">
      <div class="port-card-lbl">${isAr ? 'ألفا ضد SPY' : 'Alpha vs SPY'}</div>
      <div class="port-card-val ${alphaClass}">${fmtPct(avgAlpha)}</div>
      <div class="port-card-sub">${avgAlpha > 0 ? (isAr ? '✓ يتفوق' : '✓ outperforming') : avgAlpha < 0 ? (isAr ? '✗ يتأخر' : '✗ underperforming') : '—'}</div>
    </div>
  </div>`;

  const tableHdr = `<div class="port-thead" style="grid-template-columns:1.4fr 0.9fr 0.7fr 0.9fr 0.9fr 0.7fr 40px">
    <div class="port-tcell port-tcell-ticker">${isAr ? 'السهم / المصدر' : 'Stock / Source'}</div>
    <div class="port-tcell">${isAr ? 'دخول / حالياً' : 'Entry / Now'}</div>
    <div class="port-tcell">${isAr ? 'أيام' : 'Days'}</div>
    <div class="port-tcell port-tcell-pl">${isAr ? 'العائد' : 'Return'}</div>
    <div class="port-tcell port-tcell-pl">${isAr ? 'SPY' : 'SPY'}</div>
    <div class="port-tcell port-tcell-pl">${isAr ? 'ألفا' : 'Alpha'}</div>
    <div class="port-tcell port-tcell-actions"></div>
  </div>`;

  const rowsHtml = rows.map(r => {
    const aClass = cls(r.alpha_pct);
    const isOpen = !r.closed;
    return `<div class="port-trow" style="grid-template-columns:1.4fr 0.9fr 0.7fr 0.9fr 0.9fr 0.7fr 40px">
      <div class="port-tcell port-tcell-ticker" onclick="onTickerClick('${r.ticker}'); this.closest('.port-modal-bg').remove();">
        <div class="port-tt-name">${r.ticker}</div>
        <div class="port-tt-sym" style="text-transform:none;letter-spacing:0">${r.source || 'manual'}${r.confidence ? ` · ${Math.round(r.confidence)}%` : ''}</div>
      </div>
      <div class="port-tcell">
        <div class="port-tc-main">$${(r.current_price ?? 0).toFixed(2)}</div>
        <div class="port-tc-sub">${isAr?'دخول:':'in:'} $${r.entry_price.toFixed(2)}</div>
      </div>
      <div class="port-tcell">
        <div class="port-tc-main">${r.days_held}d</div>
        <div class="port-tc-sub">${new Date(r.saved_at).toLocaleDateString()}</div>
      </div>
      <div class="port-tcell port-tcell-pl">
        <div class="port-tc-main ${cls(r.ticker_return_pct)}">${fmtPct(r.ticker_return_pct)}</div>
      </div>
      <div class="port-tcell port-tcell-pl">
        <div class="port-tc-main ${cls(r.spy_return_pct)}" style="font-size:11.5px">${fmtPct(r.spy_return_pct)}</div>
      </div>
      <div class="port-tcell port-tcell-pl">
        <div class="port-tc-main ${aClass}">${fmtPct(r.alpha_pct)}</div>
        <div class="port-tc-sub ${aClass}">${r.is_winner ? '✓' : r.is_winner === false ? '✗' : '—'}</div>
      </div>
      <div class="port-tcell port-tcell-actions">
        ${isOpen ? `<button class="port-row-del" title="${isAr?'حذف':'Delete'}" onclick="event.stopPropagation(); if(confirm('${isAr?'حذف':'Delete'} ${r.ticker}?')) { deleteJournalEntry(${r.id}).then(()=>renderJournalBody()); }">✕</button>` : ''}
      </div>
    </div>`;
  }).join('');

  const note = closedRows.length < 5
    ? `<div style="background:rgba(245,158,11,.08);border:1px dashed rgba(245,158,11,.4);border-radius:8px;padding:10px 14px;font-size:11.5px;color:var(--text2);line-height:1.5">
        💡 <strong>${isAr?'تنبيه إحصائي:':'Statistical note:'}</strong> ${isAr?'تحتاج إلى ما لا يقل عن 20 توصية وثلاثة أشهر من البيانات للتمييز بين المهارة والصدفة. حالياً:':'You need at least 20 picks and 3+ months of data to distinguish skill from luck. Right now:'} ${closedRows.length} ${isAr?'توصيات':'picks'}.
      </div>`
    : '';

  el.innerHTML = cards + note + `<div class="port-table">${tableHdr}${rowsHtml}</div>`;
}

// ── KNOWLEDGE GRAPH (Phase 3) ──────────────────────────────────────────────
// Session cache keyed by `${ticker}:${type}` to avoid hammering Supabase
const _kgCache = new Map();
const KG_CACHE_TTL_MS = 10 * 60 * 1000;

// Render the knowledge graph panel for a ticker on the dashboard
async function renderKgPanel(ticker) {
  const el = document.getElementById('kgPanel');
  if (!el) return;
  if (!supabaseClient) {
    el.innerHTML = `<div class="news-loading">Supabase not connected</div>`;
    return;
  }
  const isAr = lang === 'ar';
  try {
    const result = await tool_getRelatedCompanies({ ticker, type: 'all' });
    if (result.error || !result.relationships?.length) {
      const needsAuth = !isAiReady();
      el.innerHTML = `<div class="kg-empty">
        <div style="font-size:12px;color:var(--text3);margin-bottom:10px">${result.error ? '⚠️ '+result.error : (isAr?'لا توجد علاقات مسجلة لهذا الرمز بعد':'No relationships recorded for this ticker yet.')}</div>
        ${result.error || needsAuth ? '' : `<button class="kg-seed-btn" onclick="seedThisTicker('${ticker}')">${isAr?'🔄 توليد العلاقات الآن (~5 ثواني)':'🔄 Generate Relationships Now (~5s)'}</button>`}
        ${needsAuth ? `<div style="font-size:11px;color:var(--text3)">${isAr?'سجّل الدخول لتوليد العلاقات':'Sign in to generate relationships'}</div>` : ''}
      </div>`;
      return;
    }
    const TYPE_META = {
      competitor:    { icon: '⚔️', label: isAr?'منافس':'Competitor',     color: '#ef4444' },
      supplier:      { icon: '📦', label: isAr?'مورد':'Supplier',         color: '#3b82f6' },
      customer:      { icon: '🛒', label: isAr?'عميل':'Customer',         color: '#10b981' },
      partner:       { icon: '🤝', label: isAr?'شريك':'Partner',          color: '#a855f7' },
      co_dependent:  { icon: '🔗', label: isAr?'مترابط':'Co-Dependent',  color: '#f59e0b' },
      thematic_peer: { icon: '🌊', label: isAr?'موجة مشتركة':'Thematic Peer', color: '#06b6d4' }
    };
    const order = ['competitor','supplier','customer','partner','co_dependent','thematic_peer'];
    const sections = order.filter(t => result.by_type[t]?.length).map(t => {
      const meta = TYPE_META[t];
      const chips = result.by_type[t].map(r => {
        const scoreColor = r.overall_score != null ? sCol(r.overall_score) : 'var(--text3)';
        const scoreText = r.overall_score != null ? r.overall_score.toFixed(1) : '—';
        const evidence = (r.evidence || '').replace(/"/g, '&quot;');
        return `<div class="kg-chip" onclick="onTickerClick('${r.ticker}')" title="${evidence} · Strength: ${r.strength}/10">
          <span class="kg-chip-ticker">${r.ticker}</span>
          <span class="kg-chip-score" style="color:${scoreColor}">${scoreText}</span>
          <span class="kg-chip-strength" title="Strength ${r.strength}/10">${'●'.repeat(Math.round(r.strength/2))}</span>
        </div>`;
      }).join('');
      return `<div class="kg-section">
        <div class="kg-sec-hdr" style="border-left:3px solid ${meta.color}">
          <span class="kg-sec-icon">${meta.icon}</span>
          <span class="kg-sec-label">${meta.label}</span>
          <span class="kg-sec-count">${result.by_type[t].length}</span>
        </div>
        <div class="kg-chips">${chips}</div>
      </div>`;
    }).join('');
    el.innerHTML = `<div class="kg-grid">${sections}</div>
      <div class="kg-footer">${isAr?'مرّر فوق الشريحة لمعرفة سبب العلاقة':'Hover a chip to see the relationship evidence'}</div>`;
  } catch (e) {
    el.innerHTML = `<div class="news-loading">⚠️ ${e.message}</div>`;
  }
}

// Called from the dashboard "Seed Now" button — seeds one ticker and re-renders the panel
async function seedThisTicker(ticker) {
  const el = document.getElementById('kgPanel');
  if (el) el.innerHTML = `<div class="news-loading">🔄 ${lang==='ar'?'جاري توليد العلاقات لـ':'Generating relationships for'} ${ticker}...</div>`;
  await seedTickers(ticker);
  _kgCache.clear();
  await renderKgPanel(ticker);
}

// Quick console helpers
async function viewGraph(ticker) {
  const r = await tool_getRelatedCompanies({ ticker, type: 'all' });
  if (r.error) return console.error(r.error);
  if (!r.count) return console.log(`%c${ticker}: no relationships in graph`, 'color:#f59e0b');
  console.log(`%c${ticker} — ${r.count} relationships:`, 'color:#3b82f6;font-weight:bold');
  Object.entries(r.by_type).forEach(([type, rels]) => {
    console.group(`${type} (${rels.length})`);
    rels.forEach(rel => console.log(`  ${rel.ticker} (${rel.sector || '?'}, score ${rel.overall_score?.toFixed(1) || '—'}, strength ${rel.strength}/10) — ${rel.evidence}`));
    console.groupEnd();
  });
  return r;
}

async function kgStats() {
  if (!supabaseClient) return console.error('Supabase not initialized');
  const { data, error } = await supabaseClient.from('ticker_relationships').select('source_ticker, type');
  if (error) return console.error(error.message);
  const sources = new Set(data.map(r => r.source_ticker));
  const byType = data.reduce((acc, r) => { acc[r.type] = (acc[r.type] || 0) + 1; return acc; }, {});
  const stats = { total_relationships: data.length, unique_source_tickers: sources.size, by_type: byType };
  console.log('%cKnowledge Graph Stats:', 'color:#3b82f6;font-weight:bold');
  console.table(byType);
  console.log(`Total: ${data.length} relationships across ${sources.size} source tickers`);
  return stats;
}

// Seed a specific range of tickers by market-cap rank (e.g. seedKnowledgeGraphRange(101, 200))
async function seedKnowledgeGraphRange(startRank, endRank) {
  if (!isAiReady()) { console.error('Sign in first'); return; }
  const universe = TICKERS.filter(t => STOCK[t]?.marketCap)
    .sort((a, b) => (STOCK[b].marketCap || 0) - (STOCK[a].marketCap || 0));
  const targets = universe.slice(startRank - 1, endRank);
  return _seedTickers(targets, universe);
}

// Seed specific tickers by name. Examples:
//   seedTickers('BSX')              → seed Boston Scientific only
//   seedTickers(['BSX','ISRG','MDT']) → seed multiple by name
async function seedTickers(tickers) {
  if (!isAiReady()) { console.error('Sign in first'); return; }
  if (typeof tickers === 'string') tickers = [tickers];
  if (!Array.isArray(tickers) || !tickers.length) { console.error('Pass a ticker string or array'); return; }
  const universe = TICKERS.filter(t => STOCK[t]?.marketCap);
  const targets = tickers.map(t => t.toUpperCase().trim()).filter(t => TICKERS.includes(t));
  const skipped = tickers.filter(t => !TICKERS.includes(t.toUpperCase().trim()));
  if (skipped.length) console.warn(`[KG] Skipped (not in universe): ${skipped.join(', ')}`);
  if (!targets.length) { console.error('No valid tickers to seed'); return; }
  return _seedTickers(targets, universe);
}

// Seeder. Call from console: seedKnowledgeGraph(100) → top 100 by mcap
async function seedKnowledgeGraph(topN = 100) {
  if (!TICKERS.length) { console.error('Universe not loaded'); return; }
  const universe = TICKERS.filter(t => STOCK[t]?.marketCap);
  const targets = universe
    .sort((a, b) => (STOCK[b].marketCap || 0) - (STOCK[a].marketCap || 0))
    .slice(0, topN);
  return _seedTickers(targets, universe);
}

// Shared loop used by both seedKnowledgeGraph and seedKnowledgeGraphRange
async function _seedTickers(targets, universe) {
  if (!isAiReady()) { console.error('Sign in first'); return; }
  if (!targets?.length) { console.error('No targets to seed'); return; }

  console.log(`%c[KG] Seeding ${targets.length} tickers — ~${Math.ceil(targets.length * 1.2 / 60)} min`, 'color:#3b82f6;font-weight:bold');
  let totalInserted = 0, success = 0, failed = 0;
  const startTime = Date.now();

  for (let i = 0; i < targets.length; i++) {
    const t = targets[i];
    const stk = STOCK[t] || {};
    try {
      const res = await fetch(getSeedKgUrl(), {
        method: 'POST',
        headers: getProxyHeaders(),
        body: JSON.stringify({ ticker: t, name: stk.companyName || '', sector: stk.sector || '', universe })
      });
      const data = await res.json();
      if (res.ok) {
        success++;
        totalInserted += data.inserted || 0;
        console.log(`[KG] ${i+1}/${targets.length} ${t} → ${data.inserted} relationships`);
      } else {
        failed++;
        console.warn(`[KG] ${i+1}/${targets.length} ${t} FAILED:`, data.error);
      }
    } catch (e) {
      failed++;
      console.warn(`[KG] ${i+1}/${targets.length} ${t} ERROR:`, e.message);
    }
    if (i < targets.length - 1) await new Promise(r => setTimeout(r, 200));
  }

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(0);
  console.log(`%c[KG] Done in ${elapsed}s — ${success} succeeded, ${failed} failed, ${totalInserted} relationships inserted`, 'color:#10b981;font-weight:bold');
  _kgCache.clear();
  return { success, failed, totalInserted, elapsed };
}

async function tool_getRelatedCompanies({ ticker, type }) {
  if (!ticker) return { error: 'ticker required' };
  const t = ticker.toUpperCase();
  const filterType = (type && type !== 'all') ? type.toLowerCase() : null;
  const cacheKey = `${t}:${filterType || 'all'}`;

  if (_kgCache.has(cacheKey)) {
    const cached = _kgCache.get(cacheKey);
    if (Date.now() - cached.t < KG_CACHE_TTL_MS) return cached.data;
  }

  if (!supabaseClient) return { error: 'Supabase client not initialized' };
  let query = supabaseClient.from('ticker_relationships').select('related_ticker, type, evidence, strength').eq('source_ticker', t);
  if (filterType) query = query.eq('type', filterType);
  query = query.order('strength', { ascending: false }).limit(20);

  const { data, error } = await query;
  if (error) return { error: `KG query failed: ${error.message}` };
  if (!data || !data.length) {
    const result = { ticker: t, count: 0, message: 'No relationships found in knowledge graph. Either the graph is unseeded or this ticker has no recorded relationships.' };
    _kgCache.set(cacheKey, { t: Date.now(), data: result });
    return result;
  }

  // Enrich with sector + score from local data
  const enriched = data.map(r => {
    const stk = STOCK[r.related_ticker] || {};
    const m = ANNUAL[r.related_ticker]?.length ? calcMetrics(ANNUAL[r.related_ticker]) : null;
    return {
      ticker: r.related_ticker,
      type: r.type,
      strength: r.strength,
      evidence: r.evidence,
      sector: stk.sector || null,
      overall_score: m ? calcScores(m).overall : null,
      market_cap: stk.marketCap || null
    };
  });

  const result = {
    ticker: t,
    count: enriched.length,
    by_type: enriched.reduce((acc, r) => { (acc[r.type] = acc[r.type] || []).push(r); return acc; }, {}),
    relationships: enriched
  };
  _kgCache.set(cacheKey, { t: Date.now(), data: result });
  return result;
}

// ── AGENT COUNCIL ──────────────────────────────────────────────────────────
// Three specialist agents run in parallel; a Composer synthesizes their takes.

// All specialists have access to the full tool set. Their PROMPTS define their analytical lens.
const ALL_TOOL_NAMES = ['get_ticker_snapshot','screen_by_metric','screen_by_filter','compare_tickers','get_sector_overview','get_universe_stats','screen_by_performance','get_related_companies','get_portfolio','recommend_for_profile'];

const SPECIALISTS = [
  {
    id: 'fundamental',
    icon: '🧮',
    name: 'Fundamental Analyst',
    name_ar: 'محلل أساسيات',
    color: '#3b82f6',
    tools: ALL_TOOL_NAMES,
    prompt: `You are the FUNDAMENTAL ANALYST on a 3-agent research council.

Your LENS: financial quality, valuation, profitability, balance sheet health, growth trends. Frame your take through this lens regardless of question scope.

Rules:
- ALWAYS contribute your angle. For sector/macro questions, comment on the financial QUALITY of those sectors (avg scores, profitability of top names, balance sheet strength).
- Use tools to fetch real numbers. NEVER invent figures.
- Output a concise take (40-80 words): cite P/E, margins, scores, growth rates with values. Use **bold** for numbers/tickers.
- If data is genuinely unavailable after trying, say so briefly. Don't loop on failed tool calls.
- Don't refuse — always provide some value from your lens.`
  },
  {
    id: 'market',
    icon: '📊',
    name: 'Market Watcher',
    name_ar: 'مراقب السوق',
    color: '#10b981',
    tools: ALL_TOOL_NAMES,
    prompt: `You are the MARKET WATCHER on a 3-agent research council.

Your LENS: actual price, valuation, momentum, relative positioning. You are THE PRIMARY agent for "is it expensive?", "is it cheap?", price moves, performance, and 52-week positioning.

Rules:
- For price/performance questions — USE screen_by_performance (period_days: 1=today, 7=week, 30=month, 90=quarter, 365=year).
- For "is it expensive/cheap/pricey/overvalued" questions — USE get_ticker_snapshot or compare_tickers, then ALWAYS cite BOTH the actual share price ($) AND the valuation multiple (P/E). The 'price' field in tool responses includes current price, 52w range, position in 52w range, and % from 52w high — use these for context.
- For comparisons — USE compare_tickers to get peer P/Es and prices side-by-side.
- Use tools to fetch real data. NEVER invent prices.
- For follow-ups like "is it pricey?" or "what about that one?" — look at the conversation context provided to identify which ticker(s) are being asked about. Don't loop trying to discover the ticker — extract it from context.
- Output a concise take (40-80 words): cite real $ price, P/E, and 52-week positioning when discussing valuation. Use **bold** for key figures.
- Don't loop on failed tool calls. After 1-2 successful data fetches, write your take.`
  },
  {
    id: 'macro',
    icon: '🌐',
    name: 'Macro Context',
    name_ar: 'سياق ماكرو',
    color: '#a855f7',
    tools: ALL_TOOL_NAMES,
    prompt: `You are the MACRO CONTEXT analyst on a 3-agent research council.

Your LENS: sector dynamics, peer landscape, supply-chain & competitive read-throughs, structural tailwinds/risks, causal chains across companies.

Rules:
- ALWAYS contribute your angle. For single-ticker questions, frame the macro/sector context AND check related companies via get_related_companies for causal-chain reasoning.
- For "if X happens to Y, who else gets hit?" type questions — get_related_companies is your PRIMARY tool. Walk through the supply chain (suppliers/customers) and competitive landscape.
- Use get_sector_overview for sector health and get_related_companies for company-specific read-throughs. Combine them when relevant.
- Use tools to fetch real data. NEVER invent peer relationships — if the knowledge graph is empty for a ticker, say "no recorded relationships" instead of guessing.
- Output a concise take (40-80 words): sector health, peer rank, key competitors/suppliers/customers from the KG, structural risks/tailwinds. Use **bold** for key figures and tickers.
- Don't loop on failed tool calls. After 1-2 successful data fetches, write your take.`
  }
];

const COMPOSER_PROMPT = `Compose a STRUCTURED, scannable answer from three specialist takes.

OUTPUT FORMAT (mandatory):
- Start with ### Title
- Use #### for subsections
- Use **bold** for tickers and numbers
- Use emoji section markers: 💵 📈 💰 📊 🏦 🏆 🎯 📰 ⭐ ⚠️
- Each bullet = ONE LINE
- End with "#### 🎯 Bottom Line" + 1-2 sentence verdict
- 150-250 words total
- NEVER write dense paragraphs of metrics

For COMPARE questions: header "### X vs Y", "**Winner:** X ⭐", then sections 📈 Growth · 💰 Profit · 📊 Valuation · 🏦 Leverage · 🏆 Rank with side-by-side bullets.

For SINGLE-TICKER: header "### TICKER — Analysis", then sections 💵 Price/Valuation · 📊 Quality Scores · 🏆 Sector Position · 📰 News (if any).

For SCREENERS: header + markdown table | # | Ticker | Metric | Note |.

For OPEN-ENDED: header + emoji bullets.

Pull facts ONLY from the specialist takes. Don't invent numbers. If a metric isn't in the takes, skip that bullet.

If specialists disagree, add a "**⚠️ Note:**" line.

DO NOT explain your reasoning. Output the answer directly, starting with ###.`;

const MAX_AGENT_ITERATIONS = 4;

// Filter tools by allowed names for a specialist
function getSpecialistTools(allowedNames) {
  return AGENT_TOOLS.filter(t => allowedNames.includes(t.function.name));
}

// Run one specialist agent: LLM loop with its filtered tools, returns final text + trace
async function runSpecialist(spec, userMessage, signal, onStep, chatContext = '') {
  const tools = getSpecialistTools(spec.tools);
  const userContent = chatContext
    ? `Recent conversation context (for resolving pronouns and follow-ups like "is it pricey?", "compare to those"):\n${chatContext}\n\n---\nCurrent question: ${userMessage}`
    : userMessage;
  const messages = [
    { role: 'system', content: spec.prompt },
    { role: 'user', content: userContent }
  ];
  const traceSteps = [];

  for (let iter = 0; iter < MAX_AGENT_ITERATIONS; iter++) {
    const res = await fetch(getProxyUrl(), {
      method: 'POST',
      headers: getProxyHeaders(),
      body: JSON.stringify({
        model: DEEPSEEK_MODEL,
        messages, tools, tool_choice: 'auto',
        stream: false, temperature: 0.3, max_tokens: 800
      }),
      signal
    });
    if (!res.ok) throw new Error(`${spec.name}: HTTP ${res.status}`);
    const data = await res.json();
    const msg = data.choices?.[0]?.message;
    if (!msg) throw new Error(`${spec.name}: empty response`);

    if (msg.tool_calls?.length) {
      messages.push(msg);
      for (const call of msg.tool_calls) {
        let args;
        try { args = JSON.parse(call.function.arguments); } catch { args = {}; }
        const result = await executeAgentTool(call.function.name, args);
        const resultStr = JSON.stringify(result);
        traceSteps.push({ tool: call.function.name, args, ok: !result?.error, summary: result?.error ? result.error : (result?.count != null ? `${result.count} results` : 'done') });
        onStep && onStep(traceSteps);
        messages.push({
          role: 'tool', tool_call_id: call.id,
          content: resultStr.length > 3000 ? resultStr.slice(0, 3000) + '...[truncated]' : resultStr
        });
      }
      continue;
    }

    return { text: msg.content || '', trace: traceSteps };
  }
  return { text: '(reached iteration limit)', trace: traceSteps };
}

// ── RISK GATE (Phase 2) ────────────────────────────────────────────────────
// Adapted from nexus-intel RiskStrategist — 6 checks → confidence bucket.

const TICKER_ABBREVS = new Set([
  'CEO','CFO','COO','IPO','GDP','ETF','USA','USD','EUR','GBP','JPY',
  'YOY','QOQ','TTM','FCF','EPS','ROE','ROA','PE','AI','ML','IT',
  'US','EU','UK','FY','Q1','Q2','Q3','Q4','YTD','EV','NA','DR',
  'AM','PM','EST','UTC','API','FMP','LTM','NTM','MKT','CAP','HQ',
  'INC','LLC','PLC','AG','SA','KG','SE','ALL','ANY','TOP','LOW',
  'MAX','MIN','NET','TAX','OIL','GAS','CAR','EVS','PC','TV','RD',
  'AT','ON','BY','IN','UP','OF','TO','OR','IF','NO','SO','VS',
  'EM','FX','VC','PE','RE','IT','CX','AR','VR','SG','R&D'
]);

function runRiskGate(specialistResults, userMessage) {
  let confidence = 0.85;
  const passes = [];
  const warnings = [];

  // Check 1: Entity verification (hallucination probe)
  const allText = specialistResults.map(r => r.text || '').join(' ');
  const mentionedWords = [...new Set([...allText.matchAll(/\b([A-Z]{2,5})\b/g)].map(m => m[1]))];
  const mentionedTickers = mentionedWords.filter(t => !TICKER_ABBREVS.has(t));
  const unknownTickers = mentionedTickers.filter(t => !TICKERS.includes(t));
  if (unknownTickers.length > 3) {
    confidence *= 0.85;
    warnings.push(`${unknownTickers.length} unverified tickers`);
  } else {
    passes.push('entities verified');
  }

  // Check 2: Source diversity (how many specialists used data tools)
  const specWithTools = specialistResults.filter(r => (r.trace?.length || 0) > 0).length;
  const totalToolCalls = specialistResults.reduce((s, r) => s + (r.trace?.length || 0), 0);
  const isDataQuery = /\b(top|best|worst|highest|lowest|which|show|list|find|screen|gain|loss|compare|sector)\b/i.test(userMessage);
  if (specWithTools === 0 && isDataQuery) {
    confidence *= 0.60;
    warnings.push('no data tools used');
  } else if (specWithTools < 2) {
    confidence *= 0.82;
    warnings.push('only 1 specialist sourced data');
  } else {
    passes.push(`${specWithTools}/3 sourced data`);
  }

  // Check 3: Tool success rate
  if (totalToolCalls > 0) {
    const failedCalls = specialistResults.flatMap(r => r.trace || []).filter(s => !s.ok).length;
    const failRate = failedCalls / totalToolCalls;
    if (failRate > 0.5) {
      confidence *= 0.70;
      warnings.push(`${failedCalls}/${totalToolCalls} tool calls failed`);
    } else if (failedCalls > 0) {
      confidence *= 0.92;
    } else {
      passes.push('all tools succeeded');
    }
  }

  // Check 4: Iteration limit (loop / contradiction indicator)
  const hitLimit = specialistResults.filter(r => r.text?.includes('reached iteration limit')).length;
  if (hitLimit > 0) {
    confidence *= 0.72;
    warnings.push(`${hitLimit} specialist(s) hit limit`);
  }

  // Check 5: Data freshness — find DB max year, flag lagged tickers
  let maxYear = 0;
  TICKERS.forEach(t => {
    const rows = ANNUAL[t];
    if (rows?.length) maxYear = Math.max(maxYear, rows[rows.length - 1].year);
  });
  const knownMentioned = mentionedTickers.filter(t => TICKERS.includes(t)).slice(0, 12);
  const staleCount = knownMentioned.filter(t => {
    const rows = ANNUAL[t];
    return rows?.length && rows[rows.length - 1].year < maxYear - 1;
  }).length;
  if (staleCount > 0) {
    confidence *= 0.90;
    warnings.push(`${staleCount} ticker(s) have lagged financials`);
  }

  // Check 6: Response grounding — all specialists gave substantive answers
  const thinResponses = specialistResults.filter(r => !r.text?.trim() || r.text.trim().length < 20).length;
  if (thinResponses > 1) {
    confidence *= 0.80;
    warnings.push(`${thinResponses} specialists gave thin responses`);
  } else {
    passes.push('all specialists contributed');
  }

  confidence = Math.max(0.05, Math.min(1, confidence));
  const bucket = confidence < 0.4 ? 'speculative'
               : confidence < 0.6 ? 'watching'
               : confidence < 0.8 ? 'conviction'
               : 'high_conviction';

  return { confidence: Math.round(confidence * 100), bucket, passes, warnings };
}

function buildRiskGateHtml(rg) {
  const META = {
    high_conviction: ['🔒', 'High Confidence'],
    conviction:      ['✓',  'Confident'],
    watching:        ['⚡', 'Mixed Data'],
    speculative:     ['⚠️', 'Speculative'],
  };
  const [icon, label] = META[rg.bucket] || ['?', 'Unknown'];
  const checks = [
    ...rg.passes.map(p => `<span class="rg-pass">✓ ${p}</span>`),
    ...rg.warnings.map(w => `<span class="rg-warn">⚠ ${w}</span>`)
  ].join('');
  return `<div class="rg-strip rg-${rg.bucket}">
    <span class="rg-badge">${icon} ${label}</span>
    <div class="rg-checks">${checks}</div>
    <span class="rg-pct">${rg.confidence}%</span>
  </div>`;
}

async function runCouncil(userMessage) {
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">${lang==='ar'?'🔒 سجل الدخول لاستخدام مجلس الوكلاء':'🔒 Sign in to use the agent council'}</div>`, {
      actions: [{ label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' }]
    });
    return;
  }
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  // Build council UI: header bubble + 3 specialist panels + composed answer placeholder
  const councilMsg = document.createElement('div');
  councilMsg.className = 'chat-msg agent';
  councilMsg.innerHTML = `
    <div class="chat-msg-bubble council-shell">
      <div class="council-header">
        <span class="council-icon">⚖️</span>
        <span class="council-label">${lang==='ar'?'مجلس البحث':'Research Council'}</span>
        <span class="council-status" id="councilStatus_${Date.now()}">${lang==='ar'?'يجمع البيانات...':'Gathering data...'}</span>
      </div>
      <div class="council-panels">
        ${SPECIALISTS.map(s => `
          <div class="council-panel" data-spec="${s.id}" style="--c:${s.color}">
            <div class="cp-head" onclick="this.parentElement.classList.toggle('expanded')">
              <span class="cp-icon">${s.icon}</span>
              <span class="cp-name">${lang==='ar' ? s.name_ar : s.name}</span>
              <span class="cp-status">⏳</span>
              <span class="cp-arrow">▼</span>
            </div>
            <div class="cp-trace"></div>
            <div class="cp-take"></div>
          </div>`).join('')}
      </div>
      <div class="council-divider"></div>
      <div class="council-answer">
        <div class="ca-loading">${lang==='ar'?'بانتظار التركيب...':'Awaiting synthesis...'}</div>
      </div>
    </div>`;
  document.getElementById('chatMessages')?.appendChild(councilMsg);
  scrollChatToBottom();

  const updatePanelTrace = (specId, steps) => {
    const panel = councilMsg.querySelector(`[data-spec="${specId}"] .cp-trace`);
    if (!panel) return;
    panel.innerHTML = steps.map(s => `<div class="trace-step"><span class="trace-arrow-small">→</span> <code>${s.tool}</code> <span class="trace-status">${s.ok?'✓':'❌'} ${s.summary}</span></div>`).join('');
  };

  // Build chat context from last 4 messages (excluding the current question itself)
  const recentChat = chatHistory.slice(-5, -1)
    .map(m => `${m.role === 'user' ? 'User' : 'Assistant'}: ${(m.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 300)}`)
    .filter(s => s.length > 10)
    .join('\n');

  try {
    // Fan out: run all specialists in parallel
    const results = await Promise.all(SPECIALISTS.map(spec =>
      runSpecialist(spec, userMessage, signal, steps => updatePanelTrace(spec.id, steps), recentChat)
        .then(r => {
          const panel = councilMsg.querySelector(`[data-spec="${spec.id}"]`);
          if (panel) {
            panel.querySelector('.cp-status').textContent = '✓';
            panel.querySelector('.cp-take').innerHTML = fmtAI(r.text);
          }
          return { ...r, spec };
        })
        .catch(err => {
          const panel = councilMsg.querySelector(`[data-spec="${spec.id}"]`);
          if (panel) {
            panel.querySelector('.cp-status').textContent = '❌';
            panel.querySelector('.cp-take').textContent = `Error: ${err.message}`;
          }
          return { text: `(${spec.name} failed: ${err.message})`, trace: [], spec };
        })
    ));

    // Run risk gate and inject confidence strip before the divider
    const riskResult = runRiskGate(results, userMessage);
    const riskEl = document.createElement('div');
    riskEl.innerHTML = buildRiskGateHtml(riskResult);
    const divider = councilMsg.querySelector('.council-divider');
    if (divider && riskEl.firstChild) divider.before(riskEl.firstChild);

    // Update status
    const statusEl = councilMsg.querySelector('.council-status');
    if (statusEl) statusEl.textContent = lang==='ar' ? 'يركّب الإجابة...' : 'Composing answer...';

    // Compose final answer — non-streaming (specialists already use this pattern reliably)
    const composerInput = `USER QUESTION: ${userMessage}\n\n` +
      results.map(r => `--- ${r.spec.name.toUpperCase()} ---\n${r.text}`).join('\n\n');

    const answerEl = councilMsg.querySelector('.council-answer');
    if (answerEl) answerEl.innerHTML = `<div class="ca-content"><div class="ca-loading">${lang==='ar'?'يركّب الإجابة...':'Composing...'}</div></div>`;
    const contentEl = answerEl?.querySelector('.ca-content');

    // 25s timeout — page can never hang
    const composerCtrl = new AbortController();
    const composerTimer = setTimeout(() => composerCtrl.abort(), 25000);

    let finalText = '';
    const composerStart = Date.now();
    console.log('[Composer] starting with model:', DEEPSEEK_MODEL);
    try {
      const composerRes = await fetch(getProxyUrl(), {
        method: 'POST',
        headers: getProxyHeaders(),
        body: JSON.stringify({
          model: DEEPSEEK_MODEL,
          messages: [
            { role: 'system', content: COMPOSER_PROMPT },
            { role: 'user', content: composerInput }
          ],
          stream: false, temperature: 0.4, max_tokens: 1200
        }),
        signal: composerCtrl.signal
      });
      console.log(`[Composer] HTTP ${composerRes.status} after ${Date.now()-composerStart}ms`);

      if (!composerRes.ok) {
        const errBody = await composerRes.text();
        throw new Error(`HTTP ${composerRes.status}: ${errBody.slice(0, 200)}`);
      }
      const data = await composerRes.json();
      console.log(`[Composer] parsed in ${Date.now()-composerStart}ms total`);
      const msg = data.choices?.[0]?.message;
      finalText = msg?.content || msg?.reasoning_content || '';
      if (!finalText) {
        console.warn('[Composer] empty content. Full response:', JSON.stringify(data).slice(0, 500));
      }
    } catch (e) {
      console.error(`[Composer] failed after ${Date.now()-composerStart}ms:`, e);
      if (e.name === 'AbortError') {
        if (contentEl) contentEl.innerHTML = `<div class="ca-error">⏱️ Synthesis timed out after 25s. The specialist takes above are still valid.</div>`;
      } else {
        if (contentEl) contentEl.innerHTML = `<div class="ca-error">❌ Composer error: ${e.message}</div>`;
      }
    } finally {
      clearTimeout(composerTimer);
    }

    if (finalText && contentEl) {
      contentEl.innerHTML = fmtAI(finalText);
    } else if (!finalText && contentEl && !contentEl.querySelector('.ca-error')) {
      contentEl.innerHTML = `<div class="ca-error">❌ Synthesis returned empty. Check console for details.</div>`;
      finalText = '(synthesis failed — see panels above)';
    }
    if (statusEl) statusEl.textContent = lang==='ar' ? 'تم' : 'Done';

    chatHistory.push({ role: 'agent', content: finalText, timestamp: Date.now() });
    scrollChatToBottom();
  } catch (e) {
    if (e.name !== 'AbortError') {
      const answerEl = councilMsg.querySelector('.council-answer');
      if (answerEl) answerEl.innerHTML = `<div class="ca-error">❌ ${e.message}</div>`;
    }
  }
  aiAbortController = null;
}

// Keep old name as alias for routing code
const runAgent = runCouncil;

// ── CHAT AGENT FUNCTIONS ────────────────────────────────────────────────────
async function chatAnalyzeTicker(ticker) {
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const snap = buildTickerSnapshot(ticker);
  if (!snap) {
    addChatMessage('agent', lang === 'ar' ? `⚠️ لا توجد بيانات لـ ${ticker}` : `⚠️ No data available for ${ticker}`);
    aiAbortController = null;
    return;
  }

  // Show rich data card first (always, even without auth)
  const stk = STOCK[ticker] || {};
  const scores = snap.scores;
  const scoreHtml = `
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">
      <span style="font-family:var(--display);font-size:24px;font-weight:700;color:var(--accent)">${ticker}</span>
      <span style="font-size:12px;color:var(--text3)">${stk.sector || ''} · ${stk.industry || ''}</span>
      ${stk.price ? `<span style="font-weight:700;font-size:14px;margin-left:auto">$${stk.price.toFixed(2)}</span>` : ''}
    </div>
    <div class="chat-score-strip">
      <div class="chat-score-item"><div class="sc-lbl">${lang==='ar'?'النمو':'Growth'}</div><div class="sc-num" style="color:${sCol(scores.growth)}">${scores.growth}</div></div>
      <div class="chat-score-item"><div class="sc-lbl">${lang==='ar'?'الربحية':'Profit'}</div><div class="sc-num" style="color:${sCol(scores.profitability)}">${scores.profitability}</div></div>
      <div class="chat-score-item"><div class="sc-lbl">${lang==='ar'?'الصحة':'Health'}</div><div class="sc-num" style="color:${sCol(scores.health)}">${scores.health}</div></div>
      <div class="chat-score-item"><div class="sc-lbl">${lang==='ar'?'التدفق':'CF'}</div><div class="sc-num" style="color:${sCol(scores.cashflow)}">${scores.cashflow}</div></div>
      <div class="chat-score-item" style="background:linear-gradient(135deg,rgba(59,130,246,.1),rgba(6,182,212,.08))"><div class="sc-lbl">${lang==='ar'?'الإجمالي':'Overall'}</div><div class="sc-num" style="color:${sCol(scores.overall)}">${scores.overall}/10</div></div>
    </div>`;

  addChatMessage('agent', scoreHtml, { ticker });

  // If not signed in, show score card + action buttons but skip AI
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3);margin-bottom:8px">${lang==='ar'?'🔒 سجل الدخول للحصول على تحليل AI مفصل':'🔒 Sign in to get detailed AI analysis'}</div>`, {
      actions: [
        { label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' },
        { label: lang==='ar'?'📊 لوحة البيانات':'📊 Full Dashboard', onclick: `openDashPanel('${ticker}')` }
      ]
    });
    aiAbortController = null;
    return;
  }

  // Typing indicator
  const typing = addChatTyping();

  try {
    // Core AI analysis
    const aiLength = 'standard';
    const sysPrompt = lang === 'ar'
      ? 'أنت محلل مالي ذكي. اكتب ملخصاً موجزاً (80-120 كلمة) في فقرتين: 1) التقييم مع نقاط القوة 2) المخاطر والنظرة. استخدم أرقاماً محددة. استخدم **نص عريض** للأرقام المهمة. لا عناوين.'
      : 'You are an AI financial analyst. Write a concise summary (80-120 words) in 2 short paragraphs: 1) Assessment with key strengths 2) Risks and outlook. Use specific numbers. Use **bold** for key figures and metrics. No headers or bullet points.';

    // Build peer context for richer analysis
    const peers = findPeers(ticker, 5);
    const peerCtx = peers.map(p => {
      const ps = calcScores(calcMetrics(ANNUAL[p]));
      return `${p}: Score ${ps.overall}/10 (G:${ps.growth} P:${ps.profitability}), MCap ${STOCK[p]?.marketCap?fM(STOCK[p].marketCap):'N/A'}`;
    }).join('\n');

    // Sector ranking
    const sectorPeers = TICKERS.filter(t => STOCK[t]?.sector === stk.sector && ANNUAL[t]?.length);
    const sectorScores = sectorPeers.map(t => ({ t, s: calcScores(calcMetrics(ANNUAL[t])).overall })).sort((a,b) => b.s - a.s);
    const rankInSector = sectorScores.findIndex(x => x.t === ticker) + 1;

    removeChatTyping();
    const analysisBubble = addChatStreamBubble();
    const enrichedContext = `${snap.summary}\n\nPEER COMPARISON (same sector):\n${peerCtx}\n\nSECTOR RANK: #${rankInSector} of ${sectorPeers.length} in ${stk.sector}`;
    const analysisText = await streamWithRetry([
      { role: 'system', content: sysPrompt },
      { role: 'user', content: enrichedContext }
    ], analysisBubble, signal);

    if (!analysisText.trim()) throw new Error('Model returned empty');
    agentMemory[ticker] = { analysis: analysisText, timestamp: Date.now(), scores: snap.scores };
    scheduleMemorySave();
    chatHistory.push({ role: 'agent', content: analysisText, ticker, timestamp: Date.now() });

    // News sentiment
    if (FMP_QUOTE_KEY && !signal.aborted) {
      let news = await fetchLiveNews(ticker).catch(() => []);
      if (news?.length && isAiReady()) {
        news = await analyseNewsSentiment(news, ticker);
        const bull = news.filter(n => n.sentiment === 'bullish').length;
        const bear = news.filter(n => n.sentiment === 'bearish').length;
        const sentColor = bull > bear ? 'var(--green)' : bear > bull ? 'var(--red)' : 'var(--text3)';
        const sentLabel = bull > bear ? (lang==='ar'?'إيجابي':'Bullish') : bear > bull ? (lang==='ar'?'سلبي':'Bearish') : (lang==='ar'?'محايد':'Neutral');

        let newsHtml = `<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px"><span style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3)">${lang==='ar'?'📰 الأخبار':'📰 NEWS'}</span><span style="font-size:12px;font-weight:700;color:${sentColor}">${sentLabel}</span><span style="font-size:10px;color:var(--text3)">${bull}🟢 ${news.length-bull-bear}⚪ ${bear}🔴</span></div>`;
        newsHtml += news.slice(0, 3).map(n =>
          `<div style="margin-bottom:6px;display:flex;align-items:flex-start;gap:8px"><span class="news-sentiment ${n.sentiment}" style="flex-shrink:0;margin-top:2px">${n.sentiment === 'bullish' ? '▲' : n.sentiment === 'bearish' ? '▼' : '—'}</span><div><span style="font-size:12px;font-weight:500">${n.title}</span><div style="font-size:10px;color:var(--text3)">${n.publisher || ''} · ${n.publishedDate?.slice(0,10)||''}</div></div></div>`
        ).join('');

        addChatMessage('agent', newsHtml, { ticker });
        agentMemory[ticker].news = news.slice(0, 5).map(n => n.title).join('; ');
        agentMemory[ticker].sentiment = bull - bear;
      }
    }

    // Peer comparison
    if (!signal.aborted) {
      const peers = findPeers(ticker, 4);
      if (peers.length >= 2) {
        const peerData = peers.map(t => { const ps = buildTickerSnapshot(t); return ps ? { ticker: t, ...ps.scores } : null; }).filter(Boolean);
        if (peerData.length) {
          let tblHtml = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:6px">${lang==='ar'?'🔄 مقارنة الأقران':'🔄 PEER COMPARISON'}</div>`;
          tblHtml += `<table class="agent-peer-tbl"><thead><tr><th>${lang==='ar'?'الشركة':'Company'}</th><th>${lang==='ar'?'النمو':'Growth'}</th><th>${lang==='ar'?'الربحية':'Profit'}</th><th>${lang==='ar'?'الصحة':'Health'}</th><th>${lang==='ar'?'التدفق':'CF'}</th><th>${lang==='ar'?'الإجمالي':'Overall'}</th></tr></thead><tbody>`;
          tblHtml += `<tr class="current"><td><strong>${ticker}</strong></td><td>${scores.growth}</td><td>${scores.profitability}</td><td>${scores.health}</td><td>${scores.cashflow}</td><td><strong>${scores.overall}</strong></td></tr>`;
          for (const p of peerData) {
            const oC = p.overall > scores.overall ? 'pos' : p.overall < scores.overall ? 'neg' : '';
            tblHtml += `<tr><td>${p.ticker}</td><td>${p.growth}</td><td>${p.profitability}</td><td>${p.health}</td><td>${p.cashflow}</td><td class="${oC}"><strong>${p.overall}</strong></td></tr>`;
          }
          tblHtml += '</tbody></table>';
          addChatMessage('agent', tblHtml, { ticker });
        }
      }
    }

    // Verdict
    if (!signal.aborted) {
      const concerns = [];
      const avgDte = snap.m.lev.map(l=>l.dte).filter(v=>v!=null);
      const latestDte = avgDte.length ? avgDte[avgDte.length-1] : null;
      if (latestDte != null && latestDte > 2) concerns.push(`High D/E: ${latestDte.toFixed(2)}x`);
      const fcfs = snap.m.cf.map(c=>c.fcf).filter(v=>v!=null);
      if (fcfs.length && fcfs[fcfs.length-1] < 0) concerns.push(`Negative FCF: ${fM(fcfs[fcfs.length-1])}`);

      const verdictSys = lang === 'ar'
        ? 'أنت وكيل مالي. اكتب خلاصة في 2-3 جمل. ابدأ بـ "الخلاصة:" متبوعاً بتوصيتك. كن حاسماً.'
        : 'You are a financial agent. Write a verdict in 2-3 sentences. Start with "Verdict:" followed by your recommendation. Be decisive and specific.';

      const verdictBubble = addChatStreamBubble();
      verdictBubble.style.fontWeight = '500';
      verdictBubble.style.color = 'var(--text)';
      await streamWithRetry([
        { role: 'system', content: verdictSys },
        { role: 'user', content: `${ticker} analysis: ${analysisText.slice(0, 400)}... Scores: Overall ${scores.overall}/10. ${concerns.length ? 'Concerns: ' + concerns.join('; ') : 'No major concerns.'}` }
      ], verdictBubble, signal);

      // Action buttons
      const actions = [
        { label: lang==='ar'?'📋 لوحة البيانات':'📋 Full Dashboard', onclick: `openDashPanel('${ticker}')`, primary: true },
        { label: lang==='ar'?'🔬 تحليل أعمق':'🔬 Deeper dive', onclick: `chatDeeperDive('${ticker}')` },
      ];
      const peers = findPeers(ticker, 1);
      if (peers.length) actions.push({ label: lang==='ar'?`↔️ قارن مع ${peers[0]}`:`↔️ Compare with ${peers[0]}`, onclick: `chatCompareTickers('${ticker}','${peers[0]}')` });

      const actMsg = document.createElement('div');
      actMsg.className = 'chat-msg agent';
      const actBubble = document.createElement('div');
      actBubble.className = 'chat-actions';
      actBubble.innerHTML = actions.map(a => `<button class="chat-action-btn${a.primary ? ' primary' : ''}" onclick="${a.onclick}">${a.label}</button>`).join('');
      actMsg.appendChild(actBubble);
      document.getElementById('chatMessages')?.appendChild(actMsg);
      scrollChatToBottom();
    }

  } catch(err) {
    removeChatTyping();
    if (err.name !== 'AbortError') addChatMessage('agent', `❌ ${err.message}`);
  }
  aiAbortController = null;
}

async function chatCompareTickers(a, b) {
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const snapA = buildTickerSnapshot(a), snapB = buildTickerSnapshot(b);
  if (!snapA || !snapB) { addChatMessage('agent', `⚠️ Missing data for ${!snapA ? a : b}`); aiAbortController = null; return; }

  // Show comparison scores (always, even without auth)
  const cmpHtml = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">${lang==='ar'?'↔️ مقارنة':'↔️ HEAD-TO-HEAD'}</div>
    <table class="agent-peer-tbl"><thead><tr><th></th><th>${lang==='ar'?'النمو':'Growth'}</th><th>${lang==='ar'?'الربحية':'Profit'}</th><th>${lang==='ar'?'الصحة':'Health'}</th><th>${lang==='ar'?'التدفق':'CF'}</th><th>${lang==='ar'?'الإجمالي':'Overall'}</th></tr></thead><tbody>
    <tr class="current"><td><strong>${a}</strong></td><td>${snapA.scores.growth}</td><td>${snapA.scores.profitability}</td><td>${snapA.scores.health}</td><td>${snapA.scores.cashflow}</td><td><strong>${snapA.scores.overall}</strong></td></tr>
    <tr><td><strong>${b}</strong></td><td>${snapB.scores.growth}</td><td>${snapB.scores.profitability}</td><td>${snapB.scores.health}</td><td>${snapB.scores.cashflow}</td><td><strong>${snapB.scores.overall}</strong></td></tr>
    </tbody></table>`;
  addChatMessage('agent', cmpHtml);

  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">${lang==='ar'?'🔒 سجل الدخول للحصول على تحليل AI مقارن':'🔒 Sign in for AI-powered comparison analysis'}</div>`, {
      actions: [
        { label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' },
        { label: `📋 ${a} Dashboard`, onclick: `openDashPanel('${a}')` },
        { label: `📋 ${b} Dashboard`, onclick: `openDashPanel('${b}')` }
      ]
    });
    aiAbortController = null;
    return;
  }

  const sysP = lang === 'ar'
    ? `أنت محلل مالي. قارن بين ${a} و ${b} بإيجاز (60-100 كلمة): الفروقات الرئيسية وأيهما أقوى. استخدم البيانات المتوفرة فقط. استخدم **عريض** للأرقام. لا عناوين.`
    : `You are a financial analyst. Compare ${a} vs ${b} concisely (60-100 words): key differences and which is stronger. Use ONLY the provided data. Use **bold** for key numbers. No headers.`;

  // Build sector context for richer comparison
  const sameSector = STOCK[a]?.sector === STOCK[b]?.sector;
  let sectorCtx = '';
  if (sameSector && STOCK[a]?.sector) {
    const sec = STOCK[a].sector;
    const peers = TICKERS.filter(t => STOCK[t]?.sector === sec && ANNUAL[t]?.length).map(t => ({ t, s: calcScores(calcMetrics(ANNUAL[t])).overall })).sort((a,b) => b.s - a.s);
    const rankA = peers.findIndex(x => x.t === a) + 1;
    const rankB = peers.findIndex(x => x.t === b) + 1;
    sectorCtx = `\n\nSECTOR (${sec}, ${peers.length} stocks): ${a} ranks #${rankA}, ${b} ranks #${rankB}. Sector avg score: ${(peers.reduce((s,x) => s+x.s, 0)/peers.length).toFixed(1)}/10`;
  }

  const bubble = addChatStreamBubble();
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `${a}:\n${snapA.summary}\n\n${b}:\n${snapB.summary}${sectorCtx}` }
    ], bubble, signal);

    const actMsg = document.createElement('div');
    actMsg.className = 'chat-msg agent';
    const actDiv = document.createElement('div');
    actDiv.className = 'chat-actions';
    actDiv.innerHTML = `<button class="chat-action-btn primary" onclick="openDashPanel('${a}')">📋 ${a} Dashboard</button><button class="chat-action-btn primary" onclick="openDashPanel('${b}')">📋 ${b} Dashboard</button>`;
    actMsg.appendChild(actDiv);
    document.getElementById('chatMessages')?.appendChild(actMsg);
    scrollChatToBottom();
  } catch(e) { if (e.name !== 'AbortError') addChatMessage('agent', `❌ ${e.message}`); }
  aiAbortController = null;
}

async function chatDeeperDive(ticker) {
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">${lang==='ar'?'🔒 سجل الدخول للتحليل المتقدم':'🔒 Sign in for deep AI analysis'}</div>`, {
      actions: [{ label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' }]
    });
    return;
  }
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const snap = buildTickerSnapshot(ticker);
  if (!snap) { aiAbortController = null; return; }

  const sysP = lang === 'ar'
    ? 'أنت محلل مالي متقدم. اختر أهم اتجاه واحد وحلله بعمق (80-120 كلمة). استخدم **عريض** للأرقام. لا عناوين.'
    : 'You are an advanced analyst. Pick the single most interesting trend and analyse it deeply (80-120 words). Use **bold** for key numbers. No headers.';

  const priorAnalysis = agentMemory[ticker]?.analysis || '';
  const bubble = addChatStreamBubble();
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `${snap.summary}\n\nPrior analysis: ${priorAnalysis.slice(0, 300)}...\n\nDig deeper — find something the prior analysis missed.` }
    ], bubble, signal);
  } catch(e) { if (e.name !== 'AbortError') addChatMessage('agent', `❌ ${e.message}`); }
  aiAbortController = null;
}

async function chatFreeQuestion(text) {
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">${lang==='ar'?'🔒 سجل الدخول لطرح أسئلة حرة على AI':'🔒 Sign in to ask free-form AI questions'}</div>`, {
      actions: [{ label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' }]
    });
    return;
  }
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const smartContext = buildSmartContext(text);
  const recentChat = chatHistory.slice(-4).map(m => `${m.role}: ${m.content.slice(0,150)}`).join('\n');

  const sysP = lang === 'ar'
    ? 'أنت محلل مالي خبير لديه بيانات حقيقية ومحدثة. استخدم البيانات المتوفرة بدقة. لا تخترع أرقاماً. إذا كانت البيانات المطلوبة غير متوفرة، قل ذلك. أجب في 80-150 كلمة. استخدم **عريض** للأرقام والشركات. لا عناوين.'
    : 'You are an expert financial analyst with REAL, current data provided in context. Use the provided data precisely. NEVER invent numbers — only cite figures from the context below. If specific data is not in context, say so. Answer in 80-150 words. Use **bold** for numbers and tickers. No headers.';

  const bubble = addChatStreamBubble();
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `=== AVAILABLE DATA ===\n${smartContext}\n\n${recentChat ? '=== RECENT CHAT ===\n' + recentChat + '\n\n' : ''}=== USER QUESTION ===\n${text}` }
    ], bubble, signal);
  } catch(e) { if (e.name !== 'AbortError') addChatMessage('agent', `❌ ${e.message}`); }
  aiAbortController = null;
}

function renderRankTable(isAr, results, title, icon, fmt) {
  if (!results.length) { addChatMessage('agent', isAr ? 'لم يتم العثور على نتائج.' : 'No results found.'); return; }
  let html = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">${icon} ${title} (${results.length})</div>`;
  html += `<table class="agent-peer-tbl"><thead><tr><th>#</th><th>${isAr?'الرمز':'Ticker'}</th><th>${isAr?'القطاع':'Sector'}</th><th>${isAr?'القيمة':'Value'}</th><th>${isAr?'الدرجة':'Score'}</th></tr></thead><tbody>`;
  results.forEach((r, i) => {
    html += `<tr style="cursor:pointer" onclick="chatAnalyzeTicker('${r.ticker}')"><td>${i+1}</td><td><strong>${r.ticker}</strong></td><td style="font-size:11px;color:var(--text3)">${STOCK[r.ticker]?.sector||''}</td><td style="font-weight:700">${fmt(r.val)}</td><td style="color:${sCol(r.scores.overall)}">${r.scores.overall}/10</td></tr>`;
  });
  html += '</tbody></table>';
  addChatMessage('agent', html);
}

function chatScreener(text) {
  const lower = text.toLowerCase();
  const isAr = lang === 'ar';
  let filtered = TICKERS.filter(t => ANNUAL[t]?.length);
  let label = isAr ? 'نتائج البحث' : 'Search Results';

  // Ranking-by-metric: market cap, revenue, EPS, P/E, ROE
  const wantsTop = lower.match(/top|biggest|largest|highest|أكبر|أعلى/);
  const wantsBottom = lower.match(/lowest|smallest|bottom|cheapest|أصغر|أدنى/);
  const topN = parseInt(lower.match(/(?:top|bottom|biggest|largest|smallest|أكبر|أصغر)\s*(\d+)/)?.[1] || 10);

  if (lower.match(/market cap|mcap|قيمة سوقية/)) {
    let sorted = filtered.filter(t => STOCK[t]?.marketCap).map(t => ({ ticker: t, val: STOCK[t].marketCap, scores: calcScores(calcMetrics(ANNUAL[t])) }));
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    const results = sorted.slice(0, topN);
    return renderRankTable(isAr, results, isAr ? `${wantsBottom?'أصغر':'أكبر'} الشركات قيمة سوقية` : `${wantsBottom?'Smallest':'Largest'} by Market Cap`, '🏢', v => fM(v));
  }
  if (lower.match(/revenue|إيرادات|مبيعات/) && !lower.match(/growth/)) {
    let sorted = filtered.map(t => { const m = calcMetrics(ANNUAL[t]); return { ticker: t, val: m.latest?.revenue || 0, scores: calcScores(m) }; }).filter(x => x.val);
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    return renderRankTable(isAr, sorted.slice(0, topN), isAr ? `${wantsBottom?'أدنى':'أعلى'} الشركات إيرادات` : `${wantsBottom?'Lowest':'Highest'} by Revenue`, '💰', v => fM(v));
  }
  if (lower.match(/p\/?e|price.earning/) && !lower.match(/growth/)) {
    let sorted = filtered.filter(t => STOCK[t]?.pe > 0).map(t => ({ ticker: t, val: STOCK[t].pe, scores: calcScores(calcMetrics(ANNUAL[t])) }));
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    return renderRankTable(isAr, sorted.slice(0, topN), isAr ? `${wantsBottom?'أدنى':'أعلى'} نسب P/E` : `${wantsBottom?'Lowest':'Highest'} P/E Ratio`, '📊', v => v.toFixed(1) + 'x');
  }
  if (lower.match(/eps|earnings per share|ربحية السهم/)) {
    let sorted = filtered.map(t => { const m = calcMetrics(ANNUAL[t]); return { ticker: t, val: m.latest?.eps_diluted || 0, scores: calcScores(m) }; }).filter(x => x.val);
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    return renderRankTable(isAr, sorted.slice(0, topN), isAr ? `${wantsBottom?'أدنى':'أعلى'} EPS` : `${wantsBottom?'Lowest':'Highest'} EPS`, '💵', v => '$' + v.toFixed(2));
  }
  if (lower.match(/roe|return on equity|عائد حقوق/)) {
    let sorted = filtered.map(t => { const m = calcMetrics(ANNUAL[t]); return { ticker: t, val: m.prof?.[m.prof.length-1]?.roe || 0, scores: calcScores(m) }; }).filter(x => x.val);
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    return renderRankTable(isAr, sorted.slice(0, topN), isAr ? `${wantsBottom?'أدنى':'أعلى'} ROE` : `${wantsBottom?'Lowest':'Highest'} ROE`, '📈', v => v.toFixed(1) + '%');
  }
  if (lower.match(/net income|profit$|profits|صافي الربح/) && !lower.match(/margin/)) {
    let sorted = filtered.map(t => { const m = calcMetrics(ANNUAL[t]); return { ticker: t, val: m.latest?.net_income || 0, scores: calcScores(m) }; }).filter(x => x.val);
    sorted.sort((a, b) => wantsBottom ? a.val - b.val : b.val - a.val);
    return renderRankTable(isAr, sorted.slice(0, topN), isAr ? `${wantsBottom?'أدنى':'أعلى'} صافي ربح` : `${wantsBottom?'Lowest':'Highest'} Net Income`, '💸', v => fM(v));
  }

  if (lower.match(/undervalued|مقومة بأقل/)) {
    filtered = filtered.filter(t => { const s = STOCK[t]; return s?.pe && s.pe > 0 && s.pe < 15; });
    label = isAr ? 'أسهم مقومة بأقل من قيمتها (P/E < 15)' : 'Potentially Undervalued (P/E < 15)';
  } else if (lower.match(/growth|نمو/)) {
    filtered = filtered.filter(t => { const m = calcMetrics(ANNUAL[t]); const g = m.yoy?.[m.yoy.length-1]?.revenue_growth; return g && g > 15; });
    label = isAr ? 'أسهم نمو (نمو إيرادات > 15%)' : 'Growth Stocks (Rev Growth > 15%)';
  } else if (lower.match(/dividend|أرباح موزعة/)) {
    filtered = filtered.filter(t => STOCK[t]?.dividend && STOCK[t].dividend > 0);
    label = isAr ? 'أسهم توزيعات الأرباح' : 'Dividend Payers';
  } else if (lower.match(/low debt|ديون منخفضة/)) {
    filtered = filtered.filter(t => { const m = calcMetrics(ANNUAL[t]); const d = m.lev?.[m.lev.length-1]?.dte; return d != null && d < 0.5; });
    label = isAr ? 'ديون منخفضة (D/E < 0.5)' : 'Low Debt (D/E < 0.5)';
  } else if (lower.match(/high margin|profitable|ربحية عالية|مربحة/)) {
    filtered = filtered.filter(t => { const m = calcMetrics(ANNUAL[t]); const nm = m.prof?.[m.prof.length-1]?.net_margin; return nm && nm > 20; });
    label = isAr ? 'هوامش ربح عالية (> 20%)' : 'Highly Profitable (Net Margin > 20%)';
  } else if (lower.match(/large cap|mega cap|كبيرة|ضخمة/)) {
    filtered = filtered.filter(t => STOCK[t]?.marketCap && STOCK[t].marketCap > 100e9);
    label = isAr ? 'شركات كبيرة (> $100B)' : 'Large Cap (> $100B)';
  } else if (lower.match(/small cap|صغيرة/)) {
    filtered = filtered.filter(t => STOCK[t]?.marketCap && STOCK[t].marketCap < 10e9);
    label = isAr ? 'شركات صغيرة (< $10B)' : 'Small Cap (< $10B)';
  } else if (lower.match(/high cash|cash rich|نقدية/)) {
    filtered = filtered.filter(t => { const m = calcMetrics(ANNUAL[t]); const cf = m.latest?.free_cash_flow; return cf && cf > 0; });
    filtered = filtered.map(t => ({ t, fcf: calcMetrics(ANNUAL[t]).latest?.free_cash_flow })).sort((a,b) => b.fcf - a.fcf).map(x => x.t);
    label = isAr ? 'غنية بالتدفق النقدي' : 'Cash-Rich (by FCF)';
  } else if (lower.match(/worst|weakest|أضعف|أسوأ/) && lower.match(/stock|compan|أسهم|شرك/)) {
    filtered = filtered.map(t => ({ ticker: t, scores: calcScores(calcMetrics(ANNUAL[t])) })).sort((a, b) => a.scores.overall - b.scores.overall);
    const topN = lower.match(/(?:worst|bottom|أضعف)\s*(\d+)/)?.[1] || 10;
    const results = filtered.slice(0, parseInt(topN));
    let html = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">⚠️ ${isAr?'الأضعف أداءً':'WEAKEST PERFORMERS'} (${results.length})</div>`;
    html += `<table class="agent-peer-tbl"><thead><tr><th>${isAr?'الرمز':'Ticker'}</th><th>${isAr?'القطاع':'Sector'}</th><th>${isAr?'النمو':'Growth'}</th><th>${isAr?'الربحية':'Profit'}</th><th>${isAr?'الإجمالي':'Overall'}</th></tr></thead><tbody>`;
    results.forEach(r => { html += `<tr style="cursor:pointer" onclick="chatAnalyzeTicker('${r.ticker}')"><td><strong>${r.ticker}</strong></td><td style="font-size:11px;color:var(--text3)">${STOCK[r.ticker]?.sector||''}</td><td>${r.scores.growth}</td><td>${r.scores.profitability}</td><td style="font-weight:700;color:${sCol(r.scores.overall)}">${r.scores.overall}/10</td></tr>`; });
    html += '</tbody></table>';
    addChatMessage('agent', html);
    return;
  } else if (lower.match(/top|best|highest|أفضل|أعلى/) || (lower.match(/strongest|أقوى/) && lower.match(/stock|compan|أسهم|شرك/))) {
    // Sort by overall score
  } else if (lower.match(/sector|قطاع|strongest|weakest|أقوى|أضعف/)) {
    const sectorCounts = {};
    TICKERS.forEach(t => { const s = STOCK[t]?.sector; if (s) { if (!sectorCounts[s]) sectorCounts[s] = { count: 0, totalScore: 0 }; sectorCounts[s].count++; const scores = calcScores(calcMetrics(ANNUAL[t])); sectorCounts[s].totalScore += scores.overall; } });
    const sorted = Object.entries(sectorCounts).map(([s, d]) => ({ sector: s, count: d.count, avg: (d.totalScore / d.count).toFixed(1) })).sort((a, b) => b.avg - a.avg);
    let html = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">${isAr?'📊 تحليل القطاعات':'📊 SECTOR ANALYSIS'}</div>`;
    html += `<table class="agent-peer-tbl"><thead><tr><th>${isAr?'القطاع':'Sector'}</th><th>${isAr?'الشركات':'Companies'}</th><th>${isAr?'متوسط الدرجة':'Avg Score'}</th></tr></thead><tbody>`;
    sorted.forEach(s => { html += `<tr><td>${s.sector}</td><td>${s.count}</td><td style="font-weight:700;color:${sCol(parseFloat(s.avg))}">${s.avg}/10</td></tr>`; });
    html += '</tbody></table>';
    addChatMessage('agent', html);
    return;
  }

  // Sort by overall score descending
  filtered = filtered.map(t => ({ ticker: t, scores: calcScores(calcMetrics(ANNUAL[t])) })).sort((a, b) => b.scores.overall - a.scores.overall);

  const fallbackTopN = lower.match(/top\s*(\d+)/)?.[1] || 10;
  const results = filtered.slice(0, parseInt(fallbackTopN));

  if (!results.length) { addChatMessage('agent', isAr ? 'لم يتم العثور على نتائج.' : 'No results found.'); return; }

  let html = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">🔍 ${label} (${results.length}/${filtered.length})</div>`;
  html += `<table class="agent-peer-tbl"><thead><tr><th>${isAr?'الرمز':'Ticker'}</th><th>${isAr?'القطاع':'Sector'}</th><th>${isAr?'النمو':'Growth'}</th><th>${isAr?'الربحية':'Profit'}</th><th>${isAr?'الإجمالي':'Overall'}</th></tr></thead><tbody>`;
  results.forEach(r => {
    html += `<tr style="cursor:pointer" onclick="chatAnalyzeTicker('${r.ticker}')"><td><strong>${r.ticker}</strong></td><td style="font-size:11px;color:var(--text3)">${STOCK[r.ticker]?.sector||''}</td><td>${r.scores.growth}</td><td>${r.scores.profitability}</td><td style="font-weight:700;color:${sCol(r.scores.overall)}">${r.scores.overall}/10</td></tr>`;
  });
  html += '</tbody></table>';
  addChatMessage('agent', html);
}

// ── UNIVERSE DIGEST: Pre-computed market overview for AI context ────────────
let _universeDigestCache = null;
let _universeDigestTime = 0;
function buildUniverseDigest() {
  // Cache for 60 seconds
  if (_universeDigestCache && (Date.now() - _universeDigestTime < 60000)) return _universeDigestCache;

  const valid = TICKERS.filter(t => ANNUAL[t]?.length && STOCK[t]);
  const withScores = valid.map(t => {
    const m = calcMetrics(ANNUAL[t]);
    const s = calcScores(m);
    const stk = STOCK[t];
    return { t, m, s, stk, lat: m.latest || {} };
  });

  // Top 20 by market cap
  const topMcap = [...withScores].filter(x => x.stk.marketCap).sort((a,b) => b.stk.marketCap - a.stk.marketCap).slice(0, 20);

  // Top 10 by overall score
  const topScore = [...withScores].sort((a,b) => b.s.overall - a.s.overall).slice(0, 10);

  // Top 10 by revenue
  const topRev = [...withScores].filter(x => x.lat.revenue).sort((a,b) => b.lat.revenue - a.lat.revenue).slice(0, 10);

  // Sector breakdown
  const sectorMap = {};
  withScores.forEach(x => {
    const sec = x.stk.sector || 'Other';
    if (!sectorMap[sec]) sectorMap[sec] = { count: 0, totalScore: 0, totalMcap: 0, tickers: [] };
    sectorMap[sec].count++;
    sectorMap[sec].totalScore += x.s.overall;
    sectorMap[sec].totalMcap += x.stk.marketCap || 0;
    sectorMap[sec].tickers.push(x.t);
  });
  const sectors = Object.entries(sectorMap).map(([sec, d]) => ({
    sector: sec, count: d.count, avgScore: (d.totalScore / d.count).toFixed(1), mcap: d.totalMcap
  })).sort((a,b) => b.mcap - a.mcap);

  // Format as compact text for AI
  const digest = `MARKET UNIVERSE (${valid.length} stocks across ${sectors.length} sectors):

TOP 20 BY MARKET CAP:
${topMcap.map((x,i) => `${i+1}. ${x.t} (${x.stk.sector}): ${fM(x.stk.marketCap)}, P/E ${x.stk.pe?.toFixed(1)||'N/A'}, Score ${x.s.overall}/10`).join('\n')}

TOP 10 BY OVERALL SCORE:
${topScore.map((x,i) => `${i+1}. ${x.t} (${x.stk.sector}): ${x.s.overall}/10 [G:${x.s.growth} P:${x.s.profitability} H:${x.s.health} CF:${x.s.cashflow}]`).join('\n')}

TOP 10 BY REVENUE:
${topRev.map((x,i) => `${i+1}. ${x.t}: ${fM(x.lat.revenue)}`).join('\n')}

SECTORS (by total market cap):
${sectors.map(s => `${s.sector}: ${s.count} stocks, ${fM(s.mcap)} mcap, avg score ${s.avgScore}/10`).join('\n')}`;

  _universeDigestCache = digest;
  _universeDigestTime = Date.now();
  return digest;
}

// ── INTENT-AWARE CONTEXT BUILDER ────────────────────────────────────────────
function buildSmartContext(text) {
  const lower = text.toLowerCase();
  const parts = [buildUniverseDigest()];

  // Detect mentioned tickers and inject their full snapshots
  const mentioned = new Set();
  const upperText = text.toUpperCase();
  const words = upperText.replace(/[^A-Z0-9\s]/g, ' ').split(/\s+/);
  for (const w of words) {
    if (TICKERS.includes(w) && ANNUAL[w]?.length) mentioned.add(w);
  }
  if (mentioned.size > 0 && mentioned.size <= 5) {
    const snapshots = [...mentioned].map(t => {
      const snap = buildTickerSnapshot(t);
      return snap ? `\n--- ${t} FULL SNAPSHOT ---\n${snap.summary}` : '';
    }).join('\n');
    parts.push(`MENTIONED TICKERS (full data):${snapshots}`);
  }

  // Detect mentioned sectors and inject sector members
  const sectors = [...new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean))];
  for (const sec of sectors) {
    if (lower.includes(sec.toLowerCase())) {
      const members = TICKERS.filter(t => STOCK[t]?.sector === sec && ANNUAL[t]?.length);
      const top = members.map(t => {
        const s = calcScores(calcMetrics(ANNUAL[t]));
        return { t, score: s.overall, mcap: STOCK[t]?.marketCap || 0 };
      }).sort((a,b) => b.mcap - a.mcap).slice(0, 15);
      parts.push(`\n${sec.toUpperCase()} SECTOR (${members.length} stocks, top 15 by mcap):
${top.map(x => `${x.t}: ${fM(x.mcap)}, score ${x.score}/10`).join('\n')}`);
      break;
    }
  }

  // Detect P/E queries
  if (lower.match(/p\/?e|price.earning|أرباح/)) {
    const peData = TICKERS.filter(t => STOCK[t]?.pe > 0).map(t => ({ t, pe: STOCK[t].pe, sector: STOCK[t].sector })).sort((a,b) => b.pe - a.pe);
    parts.push(`\nP/E DATA: Highest 10: ${peData.slice(0,10).map(x => `${x.t}(${x.pe.toFixed(1)})`).join(', ')}; Lowest 10: ${peData.slice(-10).reverse().map(x => `${x.t}(${x.pe.toFixed(1)})`).join(', ')}`);
  }

  // Detect dividend queries
  if (lower.match(/dividend|yield|توزيع/)) {
    const divData = TICKERS.filter(t => STOCK[t]?.dividend > 0).map(t => ({ t, div: STOCK[t].dividend, price: STOCK[t].price, yield: STOCK[t].price ? (STOCK[t].dividend / STOCK[t].price * 100) : 0 })).sort((a,b) => b.yield - a.yield);
    parts.push(`\nDIVIDEND DATA: Top 15 yields: ${divData.slice(0,15).map(x => `${x.t}(${x.yield.toFixed(2)}%)`).join(', ')}`);
  }

  return parts.join('\n\n');
}

async function chatPerformanceScreen(text) {
  const isAr = lang === 'ar';
  const lower = text.toLowerCase();

  // Parse period — flexible: "4 days", "2 weeks", "9 months", "3 years", or shorthand "1d/5d/1m/6m/1y"
  let periodDays = 30, periodLabel = isAr ? 'شهر' : '1 Month';

  // Try "N <unit>" pattern first (handles "last 4 days", "past 2 weeks", etc.)
  const numUnitMatch = lower.match(/(\d+)\s*(day|week|month|year|يوم|أسبوع|شهر|سنة|أيام|أسابيع|أشهر|سنوات)/);
  if (numUnitMatch) {
    const n = parseInt(numUnitMatch[1]);
    const unit = numUnitMatch[2];
    if (unit.match(/day|يوم|أيام/)) { periodDays = n; periodLabel = isAr ? `${n} ${n>10?'يوم':'أيام'}` : `${n} Day${n>1?'s':''}`; }
    else if (unit.match(/week|أسبوع|أسابيع/)) { periodDays = n * 7; periodLabel = isAr ? `${n} ${n>10?'أسبوع':'أسابيع'}` : `${n} Week${n>1?'s':''}`; }
    else if (unit.match(/month|شهر|أشهر/)) { periodDays = n * 30; periodLabel = isAr ? `${n} ${n>10?'شهر':'أشهر'}` : `${n} Month${n>1?'s':''}`; }
    else if (unit.match(/year|سنة|سنوات/)) { periodDays = n * 365; periodLabel = isAr ? `${n} ${n>10?'سنة':'سنوات'}` : `${n} Year${n>1?'s':''}`; }
  }
  // Shorthand patterns
  else if (lower.match(/today|1\s*d\b|يوم/)) { periodDays = 1; periodLabel = isAr ? 'يوم واحد' : '1 Day'; }
  else if (lower.match(/this week|1\s*w\b|5\s*d\b|أسبوع/)) { periodDays = 7; periodLabel = isAr ? 'أسبوع' : '1 Week'; }
  else if (lower.match(/this year|1\s*y\b|ytd|سنة/)) { periodDays = 365; periodLabel = isAr ? 'سنة' : '1 Year'; }
  else if (lower.match(/6\s*m\b/)) { periodDays = 180; periodLabel = '6 Months'; }
  else if (lower.match(/3\s*m\b/)) { periodDays = 90; periodLabel = '3 Months'; }
  else if (lower.match(/this month|1\s*m\b|شهر/)) { periodDays = 30; periodLabel = isAr ? 'شهر' : '1 Month'; }

  // Map periodDays to FMP API range for the special 1d intraday endpoint
  const period = periodDays === 1 ? '1d' : 'historical';

  // Parse sector filter
  let sectorFilter = null;
  const sectors = [...new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean))];
  for (const s of sectors) {
    if (lower.includes(s.toLowerCase())) { sectorFilter = s; break; }
  }

  // Parse threshold (e.g., "gained 10%", "dropped 5%", "up more than 20%")
  let threshold = null, thresholdDir = null;
  const threshMatch = lower.match(/(gain|increase|up|rose|grew|ارتفع|صعد).*?(\d+)\s*%/) ||
                      lower.match(/(\d+)\s*%.*?(gain|increase|up|rise|ارتفع|صعد)/);
  const threshDropMatch = lower.match(/(drop|decrease|down|fell|lost|decline|انخفض|هبط).*?(\d+)\s*%/) ||
                          lower.match(/(\d+)\s*%.*?(drop|decrease|down|fall|decline|انخفض|هبط)/);
  if (threshMatch) {
    threshold = parseFloat(threshMatch[2] || threshMatch[1]);
    thresholdDir = 'up';
  } else if (threshDropMatch) {
    threshold = parseFloat(threshDropMatch[2] || threshDropMatch[1]);
    thresholdDir = 'down';
  }

  // Parse top/bottom N
  let topN = 10, showBottom = false;
  const topMatch = lower.match(/top\s*(\d+)/);
  const bottomMatch = lower.match(/(?:bottom|worst|lowest|أضعف|أسوأ)\s*(\d+)?/);
  if (topMatch) topN = parseInt(topMatch[1]);
  if (bottomMatch) { showBottom = true; topN = parseInt(bottomMatch[1]) || 10; }

  // Filter tickers
  let tickers = TICKERS.filter(t => ANNUAL[t]?.length && STOCK[t]?.marketCap);
  if (sectorFilter) tickers = tickers.filter(t => STOCK[t]?.sector === sectorFilter);

  addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">📈 ${isAr ? 'جاري حساب الأداء...' : 'Calculating performance...'} (${tickers.length} ${isAr?'شركة':'stocks'}, ${periodLabel}${sectorFilter ? ' · ' + sectorFilter : ''})</div>`);

  // For 1d: use existing changePct data
  let perfData;
  if (period === '1d') {
    perfData = tickers.map(t => ({
      ticker: t,
      perf: STOCK[t]?.changePct ?? 0,
      price: STOCK[t]?.price || 0
    }));
  } else {
    // Fetch historical prices in batches (use periodDays for arbitrary periods)
    // Add buffer for weekends/holidays — fetch ~50% more days so we have enough trading days
    const lookbackDays = Math.max(periodDays + Math.ceil(periodDays * 0.4), periodDays + 3);
    const fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - lookbackDays);
    const fromStr = fromDate.toISOString().slice(0, 10);

    // Calculate target trading bars: ~5 trading days per 7 calendar days
    const targetBars = Math.max(2, Math.round(periodDays * 5 / 7));

    perfData = [];
    const batchSize = 5;
    for (let i = 0; i < tickers.length; i += batchSize) {
      const batch = tickers.slice(i, i + batchSize);
      const promises = batch.map(async t => {
        try {
          const res = await fetch(`https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${t}&from=${fromStr}&apikey=${FMP_QUOTE_KEY}`);
          if (!res.ok) return { ticker: t, perf: null, price: STOCK[t]?.price || 0 };
          const data = await res.json();
          if (!Array.isArray(data) || data.length < 2) return { ticker: t, perf: null, price: STOCK[t]?.price || 0 };
          // Data is sorted newest-first. Take the bar that's exactly `targetBars - 1` positions back.
          const newest = data[0]?.close || data[0]?.price;
          const oldestIdx = Math.min(targetBars - 1, data.length - 1);
          const oldest = data[oldestIdx]?.close || data[oldestIdx]?.price;
          if (!oldest || !newest) return { ticker: t, perf: null, price: STOCK[t]?.price || 0 };
          return { ticker: t, perf: ((newest - oldest) / oldest * 100), price: newest };
        } catch { return { ticker: t, perf: null, price: STOCK[t]?.price || 0 }; }
      });
      const results = await Promise.all(promises);
      perfData.push(...results);
      if (i + batchSize < tickers.length) await new Promise(r => setTimeout(r, 300));
    }
  }

  // Filter out nulls
  perfData = perfData.filter(d => d.perf !== null);

  // Apply threshold filter
  if (threshold !== null) {
    if (thresholdDir === 'up') perfData = perfData.filter(d => d.perf >= threshold);
    else perfData = perfData.filter(d => d.perf <= -threshold);
  }

  // Sort and slice
  if (showBottom) {
    perfData.sort((a, b) => a.perf - b.perf);
  } else {
    perfData.sort((a, b) => b.perf - a.perf);
  }
  const results = perfData.slice(0, topN);

  if (!results.length) {
    addChatMessage('agent', isAr ? '❌ لم يتم العثور على نتائج.' : '❌ No stocks matched the criteria.');
    return;
  }

  // Build results table
  const titleIcon = showBottom ? '📉' : '📈';
  const titleText = threshold !== null
    ? (thresholdDir === 'up'
        ? (isAr ? `أسهم ارتفعت أكثر من ${threshold}%` : `Stocks up ≥${threshold}%`)
        : (isAr ? `أسهم انخفضت أكثر من ${threshold}%` : `Stocks down ≥${threshold}%`))
    : (showBottom
        ? (isAr ? 'الأضعف أداءً' : 'Worst Performers')
        : (isAr ? 'الأفضل أداءً' : 'Top Performers'));

  let html = `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:8px">${titleIcon} ${titleText} · ${periodLabel}${sectorFilter ? ' · ' + sectorFilter : ''} (${results.length}/${perfData.length})</div>`;
  html += `<table class="agent-peer-tbl"><thead><tr><th>${isAr?'الرمز':'Ticker'}</th><th>${isAr?'القطاع':'Sector'}</th><th>${isAr?'السعر':'Price'}</th><th>${isAr?'الأداء':'Performance'}</th></tr></thead><tbody>`;
  results.forEach(r => {
    const col = r.perf >= 0 ? 'var(--green)' : 'var(--red)';
    const sign = r.perf >= 0 ? '+' : '';
    html += `<tr style="cursor:pointer" onclick="chatAnalyzeTicker('${r.ticker}')"><td><strong>${r.ticker}</strong></td><td style="font-size:11px;color:var(--text3)">${STOCK[r.ticker]?.sector||''}</td><td>$${r.price.toFixed(2)}</td><td style="font-weight:700;color:${col}">${sign}${r.perf.toFixed(2)}%</td></tr>`;
  });
  html += '</tbody></table>';
  addChatMessage('agent', html);

  // Add context to AI chat history
  chatHistory.push({ role: 'agent', content: `Performance screen: ${titleText} (${periodLabel}): ${results.slice(0,5).map(r => `${r.ticker} ${r.perf.toFixed(1)}%`).join(', ')}`, timestamp: Date.now() });
}

// ── HEATMAP with DD-List metric filter ──────────────────────────────────────
let _heatmapMetric = 'change';   // active coloring metric
let _heatmapSector = null;        // null = all sectors; else a specific sector name
const _heatmapNewsSent = {};      // ticker → net news sentiment (-N..+N), cached per session
let _heatmapNewsLoading = false;
let _hmPricesRefreshing = false;
let _hmPricesCancel = false;
let _hmPricesUpdatedAt = 0;

// Batch-fetch live quotes for the currently-visible tickers (all, or selected sector).
// Rate-limited to respect FMP Starter's 300 req/min cap (paced to ~250/min for headroom),
// with a retry pass for stragglers. Updates STOCK[].price + changePct.
const FMP_RATE = { BATCH: 5, MIN_CYCLE_MS: 1300 }; // 5 calls per 1.3s ≈ 230/min (< 300 cap, leaves headroom)
const HM_QUOTE_TTL_MS = 60 * 1000; // skip tickers refreshed within the last 60s

async function refreshHeatmapPrices() {
  if (_hmPricesRefreshing) { _hmPricesCancel = true; return; }
  if (!FMP_QUOTE_KEY) { console.error('[Heatmap] FMP key not configured'); return; }

  const visibleAll = TICKERS.filter(t => STOCK[t]?.marketCap && (!_heatmapSector || STOCK[t]?.sector === _heatmapSector))
    .sort((a, b) => (STOCK[b]?.marketCap || 0) - (STOCK[a]?.marketCap || 0));
  // Skip ones already refreshed very recently — avoids wasting calls on re-clicks
  const now = Date.now();
  const toFetch = visibleAll.filter(t => !STOCK[t]?._priceUpdatedAt || (now - STOCK[t]._priceUpdatedAt) > HM_QUOTE_TTL_MS);
  if (!toFetch.length) { _hmPricesUpdatedAt = Date.now(); showHeatmap(); return; }

  _hmPricesRefreshing = true; _hmPricesCancel = false;
  showHeatmap(); // reveals progress + Cancel

  const total = toFetch.length;
  let done = 0, ok = 0;
  const failed = [];
  const t0 = Date.now();
  const setProg = (txt) => { const p = document.getElementById('hmPriceProgress'); if (p) p.textContent = txt; };

  // One paced pass. Each batch cycle is held to ≥ MIN_CYCLE_MS to stay under the rate cap.
  async function runPass(list, isRetry) {
    for (let i = 0; i < list.length; i += FMP_RATE.BATCH) {
      if (_hmPricesCancel) break;
      const batchStart = Date.now();
      const slice = list.slice(i, i + FMP_RATE.BATCH);
      const quotes = await Promise.all(slice.map(t => fetchLiveQuote(t).catch(() => null)));
      slice.forEach((t, idx) => {
        if (quotes[idx]) { updateStockBarWithLive(t, quotes[idx]); ok++; }
        else if (!isRetry) failed.push(t);
      });
      if (!isRetry) { done += slice.length; setProg(`${done}/${total}`); }
      else setProg(`${lang==='ar'?'إعادة محاولة':'retrying'} ${Math.min(i + FMP_RATE.BATCH, list.length)}/${list.length}`);
      // Pace: hold each cycle to the minimum so we never exceed the rate cap
      if (i + FMP_RATE.BATCH < list.length) {
        const elapsed = Date.now() - batchStart;
        if (elapsed < FMP_RATE.MIN_CYCLE_MS) await new Promise(r => setTimeout(r, FMP_RATE.MIN_CYCLE_MS - elapsed));
      }
    }
  }

  await runPass(toFetch, false);
  // Retry stragglers once (likely transient 429s), after a short cool-down
  if (failed.length && !_hmPricesCancel) {
    await new Promise(r => setTimeout(r, 2500));
    await runPass(failed, true);
  }

  _hmPricesRefreshing = false;
  _hmPricesUpdatedAt = Date.now();
  const finalOk = ok, elapsed = ((Date.now() - t0) / 1000).toFixed(0);
  console.log(`[Heatmap] Refreshed ${finalOk}/${total} quotes in ${elapsed}s (rate-limited to ~250/min)`);
  showHeatmap();
}

// Metric definitions. `get` returns a numeric value per ticker. `type` drives coloring.
function _latestAnnual(t) { const a = ANNUAL[t]; return a?.length ? a[a.length - 1] : null; }
const HEATMAP_METRICS = {
  change:     { label: 'Daily Change %', ar: 'التغير اليومي', type: 'diverging_pct', needsNews: false, get: t => STOCK[t]?.changePct, fmt: v => v == null ? '—' : (v >= 0 ? '+' : '') + v.toFixed(2) + '%' },
  mcap:       { label: 'Market Cap',     ar: 'القيمة السوقية', type: 'sequential', needsNews: false, get: t => STOCK[t]?.marketCap, fmt: v => v == null ? '—' : fM(v) },
  price:      { label: 'Price',          ar: 'السعر',         type: 'sequential', needsNews: false, get: t => STOCK[t]?.price, fmt: v => v == null ? '—' : '$' + v.toFixed(2) },
  revenue:    { label: 'Revenue',        ar: 'الإيرادات',      type: 'sequential', needsNews: false, get: t => _latestAnnual(t)?.revenue, fmt: v => v == null ? '—' : fM(v) },
  net_income: { label: 'Net Income',     ar: 'صافي الدخل',     type: 'diverging',  needsNews: false, get: t => _latestAnnual(t)?.net_income, fmt: v => v == null ? '—' : fM(v) },
  fcf:        { label: 'Free Cash Flow', ar: 'التدفق الحر',    type: 'diverging',  needsNews: false, get: t => _latestAnnual(t)?.free_cash_flow, fmt: v => v == null ? '—' : fM(v) },
  news_bull:  { label: 'Bullish News',   ar: 'أخبار إيجابية',  type: 'news_bull',  needsNews: true,  get: t => _heatmapNewsSent[t], fmt: v => v == null ? '—' : (v > 0 ? '🟢+' + v : v < 0 ? '🔴' + v : '⚪0') },
  news_bear:  { label: 'Bearish News',   ar: 'أخبار سلبية',    type: 'news_bear',  needsNews: true,  get: t => _heatmapNewsSent[t], fmt: v => v == null ? '—' : (v < 0 ? '🔴' + v : v > 0 ? '🟢+' + v : '⚪0') },
};

const HM_NEUTRAL = '#374151';
function _hmColorFor(value, metric, ctx) {
  if (value == null || !isFinite(value)) return HM_NEUTRAL;
  if (metric.type === 'diverging_pct') { // fixed % thresholds (daily change)
    if (value > 3) return '#0f766e'; if (value > 2) return '#10b981'; if (value > 1) return '#34d399';
    if (value > 0.05) return '#16654e'; if (value >= -0.05) return HM_NEUTRAL;
    if (value > -1) return '#7f1d1d'; if (value > -2) return '#dc2626'; if (value > -3) return '#ef4444'; return '#991b1b';
  }
  if (metric.type === 'diverging') { // relative to max abs in the set (net income, fcf)
    const n = ctx.maxAbs ? value / ctx.maxAbs : 0;
    if (n > 0.4) return '#0f766e'; if (n > 0.15) return '#10b981'; if (n > 0.01) return '#16654e';
    if (n >= -0.01) return HM_NEUTRAL;
    if (n > -0.15) return '#7f1d1d'; if (n > -0.4) return '#dc2626'; return '#991b1b';
  }
  if (metric.type === 'news_bull' || metric.type === 'news_bear') { // sentiment score
    if (value > 0) return value >= 2 ? '#0f766e' : '#16654e';
    if (value < 0) return value <= -2 ? '#991b1b' : '#7f1d1d';
    return HM_NEUTRAL;
  }
  // sequential — rank within set (mcap, price, revenue)
  const r = ctx.rank ?? 0.5;
  if (r > 0.85) return '#0f766e'; if (r > 0.65) return '#10b981'; if (r > 0.45) return '#16654e';
  if (r > 0.25) return '#1f3a4d'; return HM_NEUTRAL;
}

function setHeatmapMetric(m) {
  _heatmapMetric = m;
  const def = HEATMAP_METRICS[m];
  // News metrics require a selected sector (scope) + auth
  if (def?.needsNews) {
    if (!_heatmapSector) { showHeatmap(); return; } // render will show the "pick a sector" hint
    loadSectorNewsSentiment(_heatmapSector);
  }
  showHeatmap();
}
function setHeatmapSector(s) {
  _heatmapSector = s || null;
  const def = HEATMAP_METRICS[_heatmapMetric];
  if (def?.needsNews && _heatmapSector) loadSectorNewsSentiment(_heatmapSector);
  showHeatmap();
}

// Fetch + classify news for every ticker in ONE sector, build net-sentiment map. Scoped & cached.
async function loadSectorNewsSentiment(sector) {
  if (_heatmapNewsLoading) return;
  if (!isAiReady()) return;
  const members = TICKERS.filter(t => STOCK[t]?.sector === sector && STOCK[t]?.marketCap);
  const todo = members.filter(t => _heatmapNewsSent[t] === undefined);
  if (!todo.length) { showHeatmap(); return; }
  _heatmapNewsLoading = true;
  showHeatmap(); // shows the progress bar

  let done = 0;
  const BATCH = 4;
  for (let i = 0; i < todo.length; i += BATCH) {
    const slice = todo.slice(i, i + BATCH);
    await Promise.all(slice.map(async t => {
      try {
        const news = await fetchLiveNews(t);
        if (!news?.length) { _heatmapNewsSent[t] = 0; return; }
        const scored = await analyseNewsSentiment(news, t);
        _heatmapNewsSent[t] = scored.reduce((s, n) => s + (n.sentimentScore || 0), 0);
      } catch { _heatmapNewsSent[t] = 0; }
    }));
    done += slice.length;
    const prog = document.getElementById('hmNewsProgress');
    if (prog) prog.textContent = `${done}/${todo.length}`;
    if (i + BATCH < todo.length) await new Promise(r => setTimeout(r, 150));
  }
  _heatmapNewsLoading = false;
  showHeatmap();
}

function showHeatmap() {
  const isAr = lang === 'ar';
  const metricDef = HEATMAP_METRICS[_heatmapMetric] || HEATMAP_METRICS.change;

  // Build sector → stocks map (respecting the sector filter)
  const sectors = {};
  TICKERS.forEach(t => {
    const s = STOCK[t];
    if (!s?.sector || !s.marketCap) return;
    if (_heatmapSector && s.sector !== _heatmapSector) return;
    if (!sectors[s.sector]) sectors[s.sector] = [];
    sectors[s.sector].push({ ticker: t, mcap: s.marketCap, value: metricDef.get(t) });
  });

  const sortedSectors = Object.entries(sectors).sort((a, b) =>
    b[1].reduce((s, x) => s + x.mcap, 0) - a[1].reduce((s, x) => s + x.mcap, 0));
  const totalMcap = sortedSectors.reduce((s, [, arr]) => s + arr.reduce((a, x) => a + x.mcap, 0), 0);

  // Compute coloring context across ALL displayed tiles (for diverging/sequential scales)
  const allVals = sortedSectors.flatMap(([, arr]) => arr.map(x => x.value)).filter(v => v != null && isFinite(v));
  const maxAbs = allVals.length ? Math.max(...allVals.map(Math.abs)) : 0;
  const sortedAsc = [...allVals].sort((a, b) => a - b);
  const rankOf = (v) => { if (v == null || !isFinite(v) || !sortedAsc.length) return 0.5; const idx = sortedAsc.findIndex(x => x >= v); return sortedAsc.length > 1 ? idx / (sortedAsc.length - 1) : 0.5; };

  // Sector dropdown options
  const allSectorNames = [...new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean))].sort();
  const sectorOptions = `<option value="">${isAr?'كل القطاعات':'All Sectors'}</option>` +
    allSectorNames.map(s => `<option value="${s}" ${_heatmapSector === s ? 'selected' : ''}>${s}</option>`).join('');
  const metricOptions = Object.entries(HEATMAP_METRICS).map(([k, def]) =>
    `<option value="${k}" ${_heatmapMetric === k ? 'selected' : ''}>${isAr ? def.ar : def.label}</option>`).join('');

  // "Refresh Prices" affordance — how many visible tickers, and freshness
  const visibleCount = TICKERS.filter(t => STOCK[t]?.marketCap && (!_heatmapSector || STOCK[t]?.sector === _heatmapSector)).length;
  const freshnessMin = _hmPricesUpdatedAt ? Math.floor((Date.now() - _hmPricesUpdatedAt) / 60000) : null;
  const etaSec = Math.ceil(visibleCount / 230 * 60);
  const etaLabel = etaSec > 90 ? `~${Math.round(etaSec/60)} min` : `~${etaSec}s`;
  const refreshBtn = _hmPricesRefreshing
    ? `<button class="hm-refresh-btn hm-refreshing" onclick="refreshHeatmapPrices()">⏹ ${isAr?'إلغاء':'Cancel'} <span id="hmPriceProgress">0/0</span></button>`
    : `<button class="hm-refresh-btn" onclick="refreshHeatmapPrices()" title="${isAr?`جلب أسعار حية لـ ${visibleCount} سهم (~${etaLabel}، ضمن حد FMP)`:`Fetch live quotes for ${visibleCount} stocks (${etaLabel}, rate-limited to respect FMP's 300/min)`}">⟳ ${isAr?'تحديث الأسعار':'Refresh Prices'} (${visibleCount} · ${etaLabel})</button>`;
  const freshnessNote = (freshnessMin != null && !_hmPricesRefreshing)
    ? `<span class="hm-fresh-note">${isAr?'محدّث منذ':'updated'} ${freshnessMin}m ${isAr?'':'ago'}</span>` : '';

  let html = `<div class="heatmap-view">`;
  html += `<div class="hm-controls">
    <span style="font-size:14px;font-weight:700">${isAr ? '🗺️ خريطة السوق' : '🗺️ Market Heatmap'}</span>
    <div class="hm-dd-group">
      <label class="hm-dd-label">${isAr?'القطاع':'Sector'}</label>
      <select class="hm-dd" onchange="setHeatmapSector(this.value)">${sectorOptions}</select>
    </div>
    <div class="hm-dd-group">
      <label class="hm-dd-label">${isAr?'حسب':'Color by'}</label>
      <select class="hm-dd" onchange="setHeatmapMetric(this.value)">${metricOptions}</select>
    </div>
    ${refreshBtn}${freshnessNote}
    <span class="hm-size-note">${isAr?'الحجم = القيمة السوقية':'Tile size = market cap'}</span>
  </div>`;

  // News-metric guards: needs sector + auth
  if (metricDef.needsNews) {
    if (!isAiReady()) {
      html += `<div class="hm-hint">🔒 ${isAr?'سجّل الدخول لتلوين الخريطة حسب الأخبار':'Sign in to color the map by news sentiment'}</div></div>`;
      getMainTarget().innerHTML = html; return;
    }
    if (!_heatmapSector) {
      html += `<div class="hm-hint">📰 ${isAr?'اختر قطاعاً محدداً أعلاه لتفعيل تلوين الأخبار (يُصنّف أخبار القطاع فقط)':'Pick a specific sector above to enable news coloring — it classifies only that sector\'s news (~15-30s)'}</div></div>`;
      getMainTarget().innerHTML = html; return;
    }
    if (_heatmapNewsLoading) {
      html += `<div class="hm-hint">⏳ ${isAr?'يحلل أخبار قطاع':'Analyzing news for'} ${_heatmapSector}... <span id="hmNewsProgress">0/0</span></div>`;
    }
  }

  sortedSectors.forEach(([sector, stocks]) => {
    stocks.sort((a, b) => b.mcap - a.mcap);
    const sectorMcap = stocks.reduce((s, x) => s + x.mcap, 0);
    const sectorPct = totalMcap ? ((sectorMcap / totalMcap) * 100).toFixed(1) : '0';

    html += `<div class="hm-sector" style="margin-bottom:8px">`;
    html += `<div style="font-size:10px;font-weight:600;color:var(--text3);margin-bottom:3px;letter-spacing:.04em">${sector} · ${sectorPct}%</div>`;
    html += `<div class="hm-grid">`;

    stocks.forEach(s => {
      const weight = Math.max(Math.sqrt(s.mcap / sectorMcap) * 100, 28);
      const bg = _hmColorFor(s.value, metricDef, { maxAbs, rank: rankOf(s.value) });
      const subtitle = metricDef.fmt(s.value);
      html += `<div class="hm-tile" onclick="onTickerClick('${s.ticker}')" style="background:${bg};flex:${Math.max(s.mcap / sectorMcap * 10, 0.3)};min-width:${Math.min(weight, 120)}px" title="${s.ticker} · ${subtitle}">
        <div class="hm-tile-ticker">${s.ticker}</div>
        <div class="hm-tile-change">${subtitle}</div>
      </div>`;
    });

    html += `</div></div>`;
  });

  html += `</div>`;
  getMainTarget().innerHTML = html;
}

async function chatWatchlistAnalysis() {
  if (!isAiReady()) {
    addChatMessage('agent', `<div style="font-size:12px;color:var(--text3)">${lang==='ar'?'🔒 سجل الدخول لتحليل المحفظة':'🔒 Sign in for portfolio analysis'}</div>`, {
      actions: [{ label: lang==='ar'?'🔑 تسجيل الدخول':'🔑 Sign In', onclick: 'showPinModal()' }]
    });
    return;
  }
  if (!watchlist.length) {
    addChatMessage('agent', lang === 'ar' ? 'قائمة المراقبة فارغة. أضف شركات أولاً.' : 'Your watchlist is empty. Add companies first.');
    return;
  }
  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();
  const signal = aiAbortController.signal;

  const wlData = watchlist.map(t => {
    const s = buildTickerSnapshot(t);
    return s ? `${t}: Overall ${s.scores.overall}/10, Growth ${s.scores.growth}, Prof ${s.scores.profitability}. ${agentMemory[t]?.analysis?.slice(0,100)||'Not yet analyzed.'}` : null;
  }).filter(Boolean).join('\n');

  const sysP = lang === 'ar'
    ? 'أنت مدير محافظ. حلل هذه المحفظة في 3-4 فقرات: جودة المحفظة، التركيز، التوصيات. 200-250 كلمة. لا عناوين.'
    : 'You are a portfolio manager. Analyze this watchlist in 3-4 paragraphs: portfolio quality, concentration risks, recommendations. 200-250 words. No headers.';

  addChatMessage('agent', `<div style="font-size:11px;font-weight:700;letter-spacing:.06em;color:var(--text3);margin-bottom:4px">${lang==='ar'?'📋 تحليل المحفظة':'📋 PORTFOLIO ANALYSIS'}</div><div style="font-size:12px;color:var(--text2)">${watchlist.join(', ')} (${watchlist.length} ${lang==='ar'?'شركات':'companies'})</div>`);

  const bubble = addChatStreamBubble();
  try {
    await streamWithRetry([
      { role: 'system', content: sysP },
      { role: 'user', content: `Watchlist:\n${wlData}` }
    ], bubble, signal);
  } catch(e) { if (e.name !== 'AbortError') addChatMessage('agent', `❌ ${e.message}`); }
  aiAbortController = null;
}

// ── DASHBOARD PANEL ─────────────────────────────────────────────────────────
function openDashPanel(ticker) {
  const panel = document.getElementById('dashPanel');
  const mainArea = document.getElementById('mainArea');
  const chatView = document.getElementById('chatView');
  if (!panel || !mainArea) return;

  panel.classList.remove('hidden');
  mainArea.classList.add('dash-open');
  if (chatView) chatView.classList.remove('hidden');

  const closeLabel = document.getElementById('dashCloseLabel');
  if (closeLabel) closeLabel.textContent = lang === 'ar' ? 'العودة للمحادثة' : 'Back to Chat';

  loadTickerIntoDash(ticker);
}

function closeDashPanel() {
  const panel = document.getElementById('dashPanel');
  const mainArea = document.getElementById('mainArea');
  if (panel) panel.classList.add('hidden');
  if (mainArea) mainArea.classList.remove('dash-open');
}

function onTickerClick(ticker) {
  // Sidebar clicks open the full dashboard directly
  const chatView = document.getElementById('chatView');
  const welcomeScreen = document.getElementById('welcomeScreen');
  if (chatView) chatView.classList.add('hidden');
  if (welcomeScreen) welcomeScreen.classList.add('hidden');
  const dp = document.getElementById('dashPanel');
  if (dp) dp.classList.add('hidden');
  document.getElementById('mainArea')?.classList.remove('dash-open');
  loadTicker(ticker);
}

function backToChat() {
  const chatView = document.getElementById('chatView');
  if (chatView) chatView.classList.remove('hidden');
  const dashFullView = document.getElementById('dashFullView');
  if (dashFullView) dashFullView.classList.add('hidden');
  const dashPanel = document.getElementById('dashPanel');
  if (dashPanel) dashPanel.classList.add('hidden');
  const ws = document.getElementById('welcomeScreen');
  if (ws) ws.classList.add('hidden');
  document.getElementById('mainArea')?.classList.remove('dash-open');
}

function getMainTarget() {
  const mainArea = document.getElementById('mainArea');
  let dfv = document.getElementById('dashFullView');
  if (!dfv) {
    dfv = document.createElement('div');
    dfv.id = 'dashFullView';
    mainArea.appendChild(dfv);
  }
  dfv.classList.remove('hidden');
  const cv = document.getElementById('chatView');
  if (cv) cv.classList.add('hidden');
  const ws = document.getElementById('welcomeScreen');
  if (ws) ws.classList.add('hidden');
  const dp = document.getElementById('dashPanel');
  if (dp) dp.classList.add('hidden');
  mainArea.classList.remove('dash-open');
  return dfv;
}

// ── INDUSTRY DRILL-DOWN ──────────────────────────────────────────────────────
function renderIndustryDashboard(industry, sector) {
  const tickers = TICKERS.filter(t => STOCK[t]?.industry === industry && ANNUAL[t]?.length);
  if (!tickers.length) return;

  getMainTarget().innerHTML = `<div class="loader"><div class="spin"></div></div>`;
  const isAr = lang === 'ar';

  setTimeout(() => {
    const companies = {};
    tickers.forEach(t => {
      const m = calcMetrics(ANNUAL[t]);
      companies[t] = { m, scores: calcScores(m), stk: STOCK[t] || {} };
    });

    const avgOf = (key) => {
      const vals = tickers.map(t => companies[t].scores[key]).filter(v => v != null);
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
    };
    const sorted = [...tickers].sort((a, b) => (companies[b].scores.overall || 0) - (companies[a].scores.overall || 0));

    const statsHtml = [
      { lbl: isAr ? 'الشركات' : 'Companies', val: tickers.length },
      { lbl: isAr ? 'الدرجة' : 'Avg Score', val: avgOf('overall')?.toFixed(1) || '—', color: sCol(avgOf('overall') || 5) },
      { lbl: isAr ? 'النمو' : 'Growth', val: avgOf('growth')?.toFixed(1) || '—', color: sCol(avgOf('growth') || 5) },
      { lbl: isAr ? 'الربحية' : 'Profit', val: avgOf('profitability')?.toFixed(1) || '—', color: sCol(avgOf('profitability') || 5) },
      { lbl: isAr ? 'الصحة' : 'Health', val: avgOf('health')?.toFixed(1) || '—', color: sCol(avgOf('health') || 5) },
    ].map(s => `<div class="sec-stat"><div class="sec-stat-lbl">${s.lbl}</div><div class="sec-stat-val" ${s.color?`style="color:${s.color}"`:''} >${s.val}</div></div>`).join('');

    const rowsHtml = sorted.map((t, i) => {
      const c = companies[t];
      const lp = c.m.prof[c.m.prof.length - 1];
      const ly = c.m.yoy[c.m.yoy.length - 1];
      return `<tr class="perf-row" onclick="loadTicker('${t}')" style="cursor:pointer">
        <td style="font-weight:700">${i + 1}</td>
        <td style="font-weight:700;color:var(--accent)">${t}</td>
        <td style="text-align:center;color:${sCol(c.scores.overall)};font-weight:700">${c.scores.overall}</td>
        <td style="text-align:center" class="${pc(ly?.revenue_growth)}">${ly?.revenue_growth != null ? (ly.revenue_growth > 0 ? '+' : '') + ly.revenue_growth.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center" class="${pc(lp?.net_margin)}">${lp?.net_margin != null ? lp.net_margin.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center">${c.stk.marketCap ? fM(c.stk.marketCap) : '—'}</td>
      </tr>`;
    }).join('');

    const hmHtml = sorted.map(t => {
      const s = companies[t].scores.overall;
      const col = sCol(s);
      const alpha = Math.max(0.15, Math.min(0.5, s / 15));
      return `<div class="hm-cell" onclick="loadTicker('${t}')" style="background:${col}${Math.round(alpha * 255).toString(16).padStart(2, '0')}"><div class="hm-cell-ticker">${t}</div><div class="hm-cell-score" style="color:${col}">${s}</div></div>`;
    }).join('');

    getMainTarget().innerHTML = `
    <div class="sec-dash">
      <div class="sec-dash-hdr">
        <div class="sec-dash-title">${industry}</div>
        <div class="sec-dash-sub"><span class="tag tag-blue" style="cursor:pointer" onclick="renderSectorDashboard('${sector.replace(/'/g, "\\'")}')">${sector}</span> · ${tickers.length} ${isAr ? 'شركة' : 'companies'}</div>
      </div>
      <div class="sec-stats-row">${statsHtml}</div>
      <div class="tbl-sec">
        <div class="sec-hdr"><div class="sec-title">${isAr ? '📋 جميع الشركات' : '📋 All Companies'}</div></div>
        <div class="tbl-card"><div class="tbl-scroll"><table>
          <thead><tr><th>#</th><th>${isAr ? 'الشركة' : 'Company'}</th><th style="text-align:center">${isAr ? 'الدرجة' : 'Score'}</th><th style="text-align:center">${isAr ? 'النمو' : 'Growth'}</th><th style="text-align:center">${isAr ? 'الهامش' : 'Margin'}</th><th style="text-align:center">${isAr ? 'القيمة' : 'MCap'}</th></tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table></div></div>
      </div>
      <div class="sec-hdr"><div class="sec-title">${isAr ? '🗺️ خريطة الأداء' : '🗺️ Heatmap'}</div></div>
      <div class="sec-heatmap">${hmHtml}</div>
    </div>`;
  }, 150);
}

// ── WATCHLIST ─────────────────────────────────────────────────────────────────
let watchlist = [];
// Portfolio: [{ ticker, shares, costBasis?, addedAt }]
let portfolio = [];
// Investor profile for the AI advisor
let investorProfile = {
  riskTolerance: null,       // 'conservative' | 'moderate' | 'aggressive'
  cashAvailable: null,        // dollars to deploy
  excludedSectors: [],        // e.g. ['Energy']
  requireDividend: false,
  updatedAt: null
};

function toggleWatchlist(ticker) {
  if (watchlist.includes(ticker)) {
    watchlist = watchlist.filter(t => t !== ticker);
  } else {
    watchlist.push(ticker);
  }
  // Update the button on current dashboard
  const btn = document.getElementById('wlBtn');
  if (btn) {
    const isIn = watchlist.includes(ticker);
    btn.textContent = isIn ? (lang==='ar'?'👁️ في المراقبة':'👁️ Watching') : (lang==='ar'?'👁️ راقب':'👁️ Watch');
    btn.style.color = isIn ? 'var(--accent2)' : '';
    btn.style.borderColor = isIn ? 'var(--accent2)' : '';
  }
  renderWatchlistPanel();
  scheduleMemorySave(); // Auto-save to cloud
}

function renderWatchlistPanel() {
  const el = document.getElementById('watchlistPanel');
  if (!el) return;
  if (!watchlist.length) { el.classList.add('hidden'); return; }
  el.classList.remove('hidden');
  const isAr = lang === 'ar';
  const items = watchlist.map(t => {
    const rows = ANNUAL[t];
    if (!rows?.length) {
      // Ticker in watchlist but no data loaded yet — show placeholder
      return `<div class="wl-item" onclick="loadTicker('${t}')">
        <div class="wl-item-left">
          <span class="wl-item-ticker">${t}</span>
          <span class="wl-item-score" style="color:var(--text3)">—</span>
          <span class="wl-item-alert ok" style="opacity:.5">${isAr ? 'بدون بيانات' : 'No data'}</span>
        </div>
        <span class="wl-remove" onclick="event.stopPropagation();toggleWatchlist('${t}')">×</span>
      </div>`;
    }
    const m = calcMetrics(rows);
    const scores = calcScores(m);
    const col = sCol(scores.overall);
    const latDte = m.lev[m.lev.length - 1]?.dte;
    const latMargin = m.prof[m.prof.length - 1]?.net_margin;
    let alertHtml = `<span class="wl-item-alert ok">${isAr ? 'مستقر' : 'Stable'}</span>`;
    if (latDte != null && latDte > 2.5) alertHtml = `<span class="wl-item-alert warn">${isAr ? 'دين مرتفع' : 'High D/E'}</span>`;
    else if (latMargin != null && latMargin < 0) alertHtml = `<span class="wl-item-alert warn">${isAr ? 'خسائر' : 'Net loss'}</span>`;
    else if (scores.overall < 4) alertHtml = `<span class="wl-item-alert warn">${isAr ? 'درجة ضعيفة' : 'Low score'}</span>`;
    return `<div class="wl-item" onclick="loadTicker('${t}')">
      <div class="wl-item-left">
        <span class="wl-item-ticker">${t}</span>
        <span class="wl-item-score" style="color:${col}">${scores.overall}/10</span>
        ${alertHtml}
      </div>
      <span class="wl-remove" onclick="event.stopPropagation();toggleWatchlist('${t}')">×</span>
    </div>`;
  }).join('');
  el.innerHTML = `<div class="watchlist-hdr"><div class="watchlist-title">👁️ ${isAr ? 'قائمة المراقبة' : 'Watchlist'} (${watchlist.length})</div><button class="wl-analyze-btn" onclick="showWatchlistDashboard()">${isAr?'📊 تحليل المحفظة':'📊 Analyse Portfolio'}</button></div><div class="watchlist-items">${items}</div>`;
}

// ── WATCHLIST DASHBOARD ──────────────────────────────────────────────────────
function showWatchlistDashboard() {
  if (!watchlist.length) return;
  const isAr = lang === 'ar';

  getMainTarget().innerHTML = `<div class="loader"><div class="spin"></div></div>`;

  setTimeout(() => {
    // Compute data for all watched stocks
    const data = {};
    watchlist.forEach(t => {
      const rows = ANNUAL[t];
      if (!rows?.length) return;
      const m = calcMetrics(rows);
      data[t] = { m, scores: calcScores(m), stk: STOCK[t] || {} };
    });
    const tickers = Object.keys(data);
    if (!tickers.length) return;

    // Summary stats
    const avgScore = tickers.reduce((s, t) => s + (data[t].scores.overall || 0), 0) / tickers.length;
    const avgGrowth = (() => { const v = tickers.flatMap(t => data[t].m.yoy.map(y => y.revenue_growth)).filter(v => v != null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; })();
    const avgMargin = (() => { const v = tickers.flatMap(t => data[t].m.prof.map(p => p.net_margin)).filter(v => v != null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; })();
    const totalMcap = tickers.reduce((s, t) => s + (data[t].stk.marketCap || 0), 0);

    // Sector concentration
    const sectors = {};
    tickers.forEach(t => { const sec = data[t].stk.sector || 'Other'; sectors[sec] = (sectors[sec] || 0) + 1; });
    const sectorSorted = Object.entries(sectors).sort((a, b) => b[1] - a[1]);
    const SECTOR_COLORS = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899','#f97316','#14b8a6','#6366f1'];
    const topSectorPct = sectorSorted.length ? (sectorSorted[0][1] / tickers.length * 100) : 0;
    const diversificationScore = Math.min(10, Math.max(1, 10 - (topSectorPct - 20) / 8)).toFixed(1);

    // Concentration bar
    const concBarHtml = sectorSorted.map(([sec, cnt], i) => {
      const pct = (cnt / tickers.length * 100);
      return `<div class="wl-conc-seg" style="width:${pct}%;background:${SECTOR_COLORS[i % SECTOR_COLORS.length]}" title="${sec}: ${cnt} (${pct.toFixed(0)}%)">${pct > 12 ? sec.slice(0, 10) : ''}</div>`;
    }).join('');

    // Heat map
    const sorted = [...tickers].sort((a, b) => (data[b].scores.overall || 0) - (data[a].scores.overall || 0));
    const hmHtml = sorted.map(t => {
      const s = data[t].scores.overall;
      const col = sCol(s);
      const alpha = Math.max(0.2, Math.min(0.55, s / 14));
      const mem = agentMemory[t];
      const skepticBadge = mem?.skepticScore != null ? `<div class="wl-hm-skeptic" style="background:${mem.skepticScore >= 7 ? 'rgba(16,185,129,.2);color:var(--green)' : mem.skepticScore >= 4 ? 'rgba(245,158,11,.2);color:var(--yellow)' : 'rgba(239,68,68,.2);color:var(--red)'}">${mem.skepticScore}/10</div>` : '';
      return `<div class="wl-hm-cell" onclick="loadTicker('${t}')" style="background:${col}${Math.round(alpha * 255).toString(16).padStart(2, '0')}">
        ${skepticBadge}
        <div class="wl-hm-ticker">${t}</div>
        <div class="wl-hm-score" style="color:${col}">${s}</div>
        <div class="wl-hm-sector">${data[t].stk.sector || ''}</div>
      </div>`;
    }).join('');

    // Risks & strengths
    const risks = [], strengths = [];
    tickers.forEach(t => {
      const s = data[t].scores;
      const m = data[t].m;
      const latDte = m.lev[m.lev.length - 1]?.dte;
      const latMargin = m.prof[m.prof.length - 1]?.net_margin;
      const latFcf = m.cf[m.cf.length - 1]?.fcf;
      if (latDte != null && latDte > 2) risks.push({ ticker: t, text: isAr ? `دين مرتفع (${latDte.toFixed(1)}x)` : `High leverage (${latDte.toFixed(1)}x)`, val: latDte });
      if (latMargin != null && latMargin < 0) risks.push({ ticker: t, text: isAr ? `خسائر صافية (${latMargin.toFixed(1)}%)` : `Net loss (${latMargin.toFixed(1)}%)`, val: latMargin });
      if (s.overall < 4) risks.push({ ticker: t, text: isAr ? `درجة ضعيفة (${s.overall}/10)` : `Low score (${s.overall}/10)`, val: s.overall });
      if (s.overall >= 8) strengths.push({ ticker: t, text: isAr ? `أداء ممتاز (${s.overall}/10)` : `Excellent (${s.overall}/10)`, val: s.overall });
      if (latFcf != null && latFcf > 0 && latMargin > 15) strengths.push({ ticker: t, text: isAr ? `هامش قوي + تدفق حر` : `Strong margin + positive FCF`, val: latMargin });
    });

    const risksHtml = risks.length ? risks.slice(0, 5).map(r => `<div class="wl-risk-item"><span style="font-weight:600">${r.ticker}</span><span style="color:var(--red)">${r.text}</span></div>`).join('') : `<div class="wl-risk-item" style="color:var(--green)">${isAr ? 'لا مخاطر كبيرة مكتشفة' : 'No major risks detected'}</div>`;
    const strHtml = strengths.length ? strengths.slice(0, 5).map(s => `<div class="wl-risk-item"><span style="font-weight:600">${s.ticker}</span><span style="color:var(--green)">${s.text}</span></div>`).join('') : `<div class="wl-risk-item" style="color:var(--text3)">${isAr ? 'لا توجد نقاط قوة بارزة' : 'No standout strengths'}</div>`;

    // Ranking table
    const rankHtml = sorted.map((t, i) => {
      const d = data[t];
      const lp = d.m.prof[d.m.prof.length - 1];
      const ly = d.m.yoy[d.m.yoy.length - 1];
      const mem = agentMemory[t];
      return `<tr onclick="loadTicker('${t}')" style="cursor:pointer">
        <td style="font-weight:700">${i + 1}</td>
        <td style="font-weight:700;color:var(--accent)">${t}</td>
        <td style="text-align:center;color:${sCol(d.scores.overall)};font-weight:700">${d.scores.overall}</td>
        <td style="text-align:center" class="${pc(ly?.revenue_growth)}">${ly?.revenue_growth != null ? (ly.revenue_growth > 0 ? '+' : '') + ly.revenue_growth.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center" class="${pc(lp?.net_margin)}">${lp?.net_margin != null ? lp.net_margin.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center">${d.stk.sector || '—'}</td>
        <td style="text-align:center">${mem?.skepticScore != null ? `<span style="color:${mem.skepticScore >= 7 ? 'var(--green)' : mem.skepticScore >= 4 ? 'var(--yellow)' : 'var(--red)'};font-weight:700">${mem.skepticScore}/10</span>` : '—'}</td>
      </tr>`;
    }).join('');

    getMainTarget().innerHTML = `
    <div class="wl-dash">
      <div class="wl-dash-hdr">
        <div>
          <div class="wl-dash-title">${isAr ? '📊 لوحة المحفظة' : '📊 Portfolio Dashboard'}</div>
          <div class="wl-dash-sub">${tickers.length} ${isAr ? 'شركة في المراقبة' : 'watched companies'} · ${sectorSorted.length} ${isAr ? 'قطاعات' : 'sectors'}</div>
        </div>
        <button class="wl-analyze-btn" id="wlAiBtn" onclick="analyseWatchlistWithAI()">🤖 ${isAr ? 'تحليل ذكي للمحفظة' : 'AI Portfolio Analysis'}</button>
      </div>

      <div class="wl-summary-grid">
        <div class="wl-sum-card"><div class="wl-sum-lbl">${isAr ? 'متوسط الدرجة' : 'Avg Score'}</div><div class="wl-sum-val" style="color:${sCol(avgScore)}">${avgScore.toFixed(1)}</div><div class="wl-sum-sub">/10</div></div>
        <div class="wl-sum-card"><div class="wl-sum-lbl">${isAr ? 'التنويع' : 'Diversification'}</div><div class="wl-sum-val" style="color:${sCol(parseFloat(diversificationScore))}">${diversificationScore}</div><div class="wl-sum-sub">/10</div></div>
        <div class="wl-sum-card"><div class="wl-sum-lbl">${isAr ? 'متوسط النمو' : 'Avg Growth'}</div><div class="wl-sum-val" style="color:${avgGrowth > 0 ? 'var(--green)' : 'var(--red)'}">${avgGrowth != null ? (avgGrowth > 0 ? '+' : '') + avgGrowth.toFixed(1) + '%' : '—'}</div></div>
        <div class="wl-sum-card"><div class="wl-sum-lbl">${isAr ? 'متوسط الهامش' : 'Avg Margin'}</div><div class="wl-sum-val" style="color:${avgMargin > 10 ? 'var(--green)' : avgMargin > 0 ? 'var(--accent)' : 'var(--red)'}">${avgMargin != null ? avgMargin.toFixed(1) + '%' : '—'}</div></div>
      </div>

      <div class="sec-hdr"><div class="sec-title">${isAr ? '📊 توزيع القطاعات' : '📊 Sector Allocation'}</div></div>
      <div class="wl-conc-bar">${concBarHtml}</div>
      ${topSectorPct > 50 ? `<div style="font-size:12px;color:var(--red);margin-bottom:14px;padding:8px 12px;background:rgba(239,68,68,.06);border-radius:8px;border:1px solid rgba(239,68,68,.15)">⚠️ ${isAr ? `تركيز عالي: ${sectorSorted[0][0]} يمثل ${topSectorPct.toFixed(0)}% من محفظتك` : `High concentration: ${sectorSorted[0][0]} is ${topSectorPct.toFixed(0)}% of your portfolio`}</div>` : ''}

      <div class="sec-hdr"><div class="sec-title">${isAr ? '🗺️ خريطة الأداء' : '🗺️ Performance Heatmap'}</div></div>
      <div class="wl-heatmap">${hmHtml}</div>

      <div class="wl-risk-grid">
        <div class="wl-risk-card">
          <div class="wl-risk-title" style="color:var(--red)">⚠️ ${isAr ? 'المخاطر' : 'Risks'}</div>
          ${risksHtml}
        </div>
        <div class="wl-risk-card">
          <div class="wl-risk-title" style="color:var(--green)">💪 ${isAr ? 'نقاط القوة' : 'Strengths'}</div>
          ${strHtml}
        </div>
      </div>

      <div class="sec-hdr"><div class="sec-title">${isAr ? '🏆 ترتيب الشركات' : '🏆 Rankings'}</div></div>
      <div class="tbl-card"><div class="tbl-scroll"><table>
        <thead><tr><th>#</th><th>${isAr ? 'الشركة' : 'Ticker'}</th><th style="text-align:center">${isAr ? 'الدرجة' : 'Score'}</th><th style="text-align:center">${isAr ? 'النمو' : 'Growth'}</th><th style="text-align:center">${isAr ? 'الهامش' : 'Margin'}</th><th style="text-align:center">${isAr ? 'القطاع' : 'Sector'}</th><th style="text-align:center">${isAr ? 'ثقة المشكك' : 'Skeptic'}</th></tr></thead>
        <tbody>${rankHtml}</tbody>
      </table></div></div>

      <div class="sec-hdr" style="margin-top:18px"><div class="sec-title">🤖 ${isAr ? 'تحليل المحفظة بالذكاء الاصطناعي' : 'AI Portfolio Analysis'}</div></div>
      <div class="wl-agent-output" id="wlAgentOutput">${isAr ? 'اضغط "تحليل ذكي للمحفظة" لتفعيل التحليل الشامل.' : 'Click "AI Portfolio Analysis" above to run a comprehensive analysis of your entire watchlist.'}</div>
    </div>`;
  }, 150);
}

// AI analysis for the entire watchlist
async function analyseWatchlistWithAI() {
  if (!isAiReady()) { showPinModal(); return; }
  if (!watchlist.length) return;

  const btn = document.getElementById('wlAiBtn');
  const output = document.getElementById('wlAgentOutput');
  if (btn) { btn.disabled = true; btn.textContent = '🤖 Analysing...'; }
  if (output) output.innerHTML = `<span class="ai-loading"><span class="dot"></span><span class="dot"></span><span class="dot"></span> ${lang==='ar'?'يحلل المحفظة...':'Analysing portfolio...'}</span>`;

  // Build portfolio summary for the AI
  const snapshots = watchlist.map(t => {
    const snap = buildTickerSnapshot(t);
    if (!snap) return null;
    const mem = agentMemory[t];
    return `${t} (${snap.stk.sector||'?'}): Score ${snap.scores.overall}/10, Growth ${snap.scores.growth}, Prof ${snap.scores.profitability}, Health ${snap.scores.health}${mem?.skepticScore != null ? ', Skeptic: ' + mem.skepticScore + '/10' : ''}`;
  }).filter(Boolean).join('\n');

  const sectors = {};
  watchlist.forEach(t => { const sec = STOCK[t]?.sector || 'Other'; sectors[sec] = (sectors[sec] || 0) + 1; });
  const sectorStr = Object.entries(sectors).map(([s, c]) => `${s}: ${c} (${(c/watchlist.length*100).toFixed(0)}%)`).join(', ');

  if (aiAbortController) { try { aiAbortController.abort(); } catch(e) {} }
  aiAbortController = new AbortController();

  const sysP = lang === 'ar'
    ? 'أنت مستشار محافظ مالي. حلل هذه المحفظة بالكامل في 4 فقرات: (1) نظرة عامة على جودة المحفظة، (2) مخاطر التركيز والتنويع، (3) أقوى وأضعف المراكز، (4) توصيات محددة للتحسين. كن صريحاً ومحدداً بالأرقام. لا عناوين.'
    : 'You are a portfolio advisor. Analyse this entire portfolio in 4 paragraphs: (1) overall portfolio quality assessment, (2) concentration risks and diversification issues, (3) strongest and weakest positions with specific reasoning, (4) specific actionable recommendations to improve the portfolio. Be direct and cite specific numbers. No headers or bullets.';

  try {
    output.textContent = '';
    await streamToElement([
      { role: 'system', content: sysP },
      { role: 'user', content: `My watchlist portfolio (${watchlist.length} stocks):\n${snapshots}\n\nSector allocation: ${sectorStr}\n\nAnalyse this as a portfolio — not individual stocks. Focus on how they work together, concentration risks, and what's missing.` }
    ], output, aiAbortController.signal);
  } catch(e) {
    if (e.name !== 'AbortError') output.textContent = `Error: ${e.message}`;
  }

  if (btn) { btn.disabled = false; btn.textContent = `🤖 ${lang==='ar'?'تحليل ذكي للمحفظة':'AI Portfolio Analysis'}`; }
  aiAbortController = null;
}

// ── SECTOR COMPOSITION (treemap) ─────────────────────────────────────────────
function renderSectorComposition() {
  if (!TICKERS.length) return;
  const isAr = lang === 'ar';
  const sectors = {};
  TICKERS.forEach(t => {
    const sec = STOCK[t]?.sector || 'Other';
    if (!sectors[sec]) sectors[sec] = { count: 0, tickers: [] };
    sectors[sec].count++;
    sectors[sec].tickers.push(t);
  });

  const total = TICKERS.length;
  const sorted = Object.entries(sectors).sort((a, b) => b[1].count - a[1].count);
  const SECTOR_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#f97316', '#14b8a6', '#6366f1', '#84cc16', '#e11d48'];

  const cells = sorted.map(([sec, data], i) => {
    const pct = (data.count / total * 100);
    const width = Math.max(80, pct * 5.5);
    const col = SECTOR_COLORS[i % SECTOR_COLORS.length];
    // Compute avg score for sector
    const scores = data.tickers.map(t => {
      const rows = ANNUAL[t];
      if (!rows?.length) return null;
      return calcScores(calcMetrics(rows)).overall;
    }).filter(v => v != null);
    const avgScore = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length) : null;

    return `<div class="tm-cell" onclick="selectSector('${sec}')" style="background:${col}22;border:1px solid ${col}44;flex-basis:${width}px;flex-grow:${Math.max(1, Math.round(pct / 5))}">
      <div class="tm-cell-name" style="color:${col}">${sec}</div>
      <div class="tm-cell-cnt">${data.count} (${pct.toFixed(0)}%)</div>
      ${avgScore != null ? `<div class="tm-cell-score" style="color:${sCol(avgScore)}">${avgScore.toFixed(1)}</div>` : ''}
    </div>`;
  }).join('');

  return `<div class="sec-comp">
    <div class="sec-hdr"><div class="sec-title">${isAr ? '🗺️ توزيع القطاعات' : '🗺️ Sector Composition'}</div></div>
    <div class="treemap-grid">${cells}</div>
  </div>`;
}

// ── CHARTS ────────────────────────────────────────────────────────────────────
Chart.defaults.font.family = "'DM Sans',sans-serif";
const PALETTE = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4'];

function getChartColors() {
  return theme === 'light'
    ? { grid:'#e2e8f0', tick:'#64748b', legend:'#475569' }
    : { grid:'#1e2d4780', tick:'#475569', legend:'#94a3b8' };
}

function destroyCharts() { Object.values(charts).forEach(c=>{try{c.destroy()}catch(e){}}); charts={}; }

function mk(id, cfg) {
  const el=document.getElementById(id); if(!el)return null;
  if(charts[id])try{charts[id].destroy()}catch(e){}
  Chart.defaults.color = getChartColors().legend;
  Chart.defaults.borderColor = getChartColors().grid;
  charts[id]=new Chart(el,cfg);
  return charts[id];
}

function scOpts(extra={}) {
  const c = getChartColors();
  return { x:{ticks:{color:c.tick},grid:{color:c.grid}}, y:{ticks:{color:c.tick},grid:{color:c.grid}}, ...extra };
}

function drawCharts(ticker, rows, m) {
  const yrs = rows.map(r=>r.year);
  destroyCharts();

  // Revenue & Net Income
  mk('cRev',{type:'line',data:{labels:yrs,datasets:[
    {label:T('revenue'),   data:rows.map(r=>r.revenue),    borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,.08)',tension:.4,fill:true,pointRadius:4},
    {label:T('netIncome'), data:rows.map(r=>r.net_income), borderColor:'#10b981',backgroundColor:'rgba(16,185,129,.08)',tension:.4,fill:true,pointRadius:4}
  ]},options:{responsive:true,maintainAspectRatio:false,scales:scOpts(),plugins:{tooltip:{callbacks:{label:c=>fM(c.raw)}}}}});

  // Cash flow
  mk('cCF',{type:'line',data:{labels:yrs,datasets:[
    {label:T('ocf'),data:m.cf.map(c=>c.ocf),borderColor:'#06b6d4',backgroundColor:'rgba(6,182,212,.08)',tension:.4,fill:true},
    {label:T('fcf'),data:m.cf.map(c=>c.fcf),borderColor:'#f59e0b',backgroundColor:'rgba(245,158,11,.08)',tension:.4,fill:true}
  ]},options:{responsive:true,maintainAspectRatio:false,scales:scOpts(),plugins:{tooltip:{callbacks:{label:c=>fM(c.raw)}}}}});

  // YoY bar
  mk('cYoY',{type:'bar',data:{labels:m.yoy.map(y=>y.year),datasets:[{
    label:T('revGrowth'),data:m.yoy.map(y=>y.revenue_growth),
    backgroundColor:m.yoy.map(y=>(y.revenue_growth||0)>=0?'rgba(16,185,129,.7)':'rgba(239,68,68,.7)'),
    borderColor:m.yoy.map(y=>(y.revenue_growth||0)>=0?'#10b981':'#ef4444'),
    borderWidth:1,borderRadius:5
  }]},options:{responsive:true,maintainAspectRatio:false,scales:scOpts(),plugins:{legend:{display:false}}}});

  // Radar
  const scores = calcScores(m);
  mk('cRadar',{type:'radar',data:{
    labels:[T('growth'),T('profitability'),T('health'),T('cashflow')],
    datasets:[{label:ticker,data:[scores.growth,scores.profitability,scores.health,scores.cashflow],
      backgroundColor:'rgba(59,130,246,.15)',borderColor:'#3b82f6',pointBackgroundColor:'#3b82f6',pointRadius:4}]
  },options:{responsive:true,maintainAspectRatio:false,
    scales:{r:{min:0,max:10,ticks:{stepSize:2,color:getChartColors().tick,backdropColor:'transparent'},grid:{color:getChartColors().grid},pointLabels:{color:getChartColors().legend,font:{size:11}}}},
    plugins:{legend:{display:false}}}});
}

// ── LOAD TICKER ───────────────────────────────────────────────────────────────
let _dashTarget = null;

// ── DECISION CARD (rule-based, token-free) ───────────────────────────────────
// Verdict + "what's priced in" computed entirely from local data. No LLM calls.
let _lastVerdict = null; // {ticker, summary, confidence} — attached to journal saves

function _dbMaxYear() {
  let y = 0;
  TICKERS.forEach(t => { const a = ANNUAL[t]; if (a?.length) y = Math.max(y, a[a.length - 1].year); });
  return y;
}

// P/E standing within the ticker's sector (uses precomputed STOCK.pe — cheap)
function computeSectorPE(ticker) {
  const sec = STOCK[ticker]?.sector;
  if (!sec) return null;
  const pes = TICKERS.filter(t => STOCK[t]?.sector === sec && STOCK[t].pe > 0 && STOCK[t].pe < 200)
    .map(t => STOCK[t].pe).sort((a, b) => a - b);
  if (pes.length < 5) return null;
  const median = pes[Math.floor(pes.length / 2)];
  const pe = STOCK[ticker]?.pe;
  const pct = (pe > 0) ? Math.round(pes.filter(x => x <= pe).length / pes.length * 100) : null;
  return { median, pct, count: pes.length };
}

// Reverse DCF: what constant 10y FCF growth does the current market cap imply?
function computeImpliedGrowth(ticker, m, opts = {}) {
  const stk = STOCK[ticker] || {};
  const lat = m.latest || {};
  const mcap = stk.marketCap || (stk.price && lat.shares_diluted ? stk.price * lat.shares_diluted : null);
  const fcf = lat.free_cash_flow;
  if (!mcap || fcf == null) return null;
  if (fcf <= 0) return { negative: true, fcf, mcap };

  const N = 10, gt = 0.025;
  const pvAt = (g, r) => {
    let pv = 0, c = fcf;
    for (let i = 1; i <= N; i++) { c = c * (1 + g); pv += c / Math.pow(1 + r, i); }
    const term = (r > gt) ? (c * (1 + gt) / (r - gt)) / Math.pow(1 + r, N) : Infinity;
    return pv + term;
  };
  const solve = (r) => {
    let lo = -0.5, hi = 1.0;
    if (pvAt(lo, r) > mcap) return lo; // cheaper than even -50%/yr implies
    if (pvAt(hi, r) < mcap) return hi;
    for (let i = 0; i < 60; i++) { const mid = (lo + hi) / 2; (pvAt(mid, r) < mcap) ? lo = mid : hi = mid; }
    return (lo + hi) / 2;
  };

  // Trailing FCF CAGR over up to 4 years (needs positive endpoints)
  let trailing = null;
  const cfs = m.cf.filter(c => c.fcf != null);
  if (cfs.length >= 3) {
    const first = cfs[Math.max(0, cfs.length - 4)], last = cfs[cfs.length - 1];
    const yrs = last.year - first.year;
    if (yrs > 0 && first.fcf > 0 && last.fcf > 0) trailing = (Math.pow(last.fcf / first.fcf, 1 / yrs) - 1) * 100;
  }

  // Analyst forward revenue CAGR from latest actual to furthest future estimate
  let analyst = null;
  const est = (ESTIMATES[ticker] || []).filter(e => e.revenueAvg && new Date(e.date) > new Date())
    .sort((a, b) => a.date.localeCompare(b.date));
  if (est.length && lat.revenue > 0) {
    const far = est[est.length - 1];
    const yrs = Math.max(1, (new Date(far.date) - new Date()) / 31557600000);
    analyst = (Math.pow(far.revenueAvg / lat.revenue, 1 / yrs) - 1) * 100;
  }

  return {
    implied: solve(0.10) * 100, sens8: solve(0.08) * 100, sens12: solve(0.12) * 100,
    trailing, analyst, pFcf: mcap / fcf, fcf, mcap,
  };
}

// Rule-based verdict: quality (scores) × valuation (sector P/E percentile or P/FCF)
function computeVerdict(ticker, m, scores) {
  const isAr = lang === 'ar';
  const L = (en, ar) => isAr ? ar : en;
  const stk = STOCK[ticker] || {};
  const lat = m.latest || {};
  const secPE = computeSectorPE(ticker);
  const ig = computeImpliedGrowth(ticker, m);

  // Valuation tier
  let val = 'na', valLine = L('Valuation: not measurable (no positive earnings or FCF)', 'التقييم: غير قابل للقياس (لا أرباح أو تدفق حر موجب)');
  if (stk.pe > 0 && secPE?.pct != null) {
    val = secPE.pct <= 40 ? 'cheap' : secPE.pct <= 70 ? 'fair' : 'rich';
    valLine = L(`P/E ${stk.pe.toFixed(1)}x vs sector median ${secPE.median.toFixed(1)}x — ${secPE.pct}th percentile of ${secPE.count}`,
                `مضاعف الربح ${stk.pe.toFixed(1)}x مقابل وسيط القطاع ${secPE.median.toFixed(1)}x — المئين ${secPE.pct} من ${secPE.count}`);
  } else if (stk.pe > 0) {
    val = stk.pe <= 15 ? 'cheap' : stk.pe <= 28 ? 'fair' : 'rich';
    valLine = L(`P/E ${stk.pe.toFixed(1)}x (sector comparison unavailable)`, `مضاعف الربح ${stk.pe.toFixed(1)}x (لا تتوفر مقارنة قطاعية)`);
  } else if (ig && !ig.negative && ig.pFcf > 0) {
    val = ig.pFcf <= 15 ? 'cheap' : ig.pFcf <= 30 ? 'fair' : 'rich';
    valLine = L(`No positive earnings — P/FCF ${ig.pFcf.toFixed(1)}x`, `لا أرباح موجبة — السعر/التدفق الحر ${ig.pFcf.toFixed(1)}x`);
  }

  // Quality tier
  const q = scores.overall >= 6.5 ? 'hi' : scores.overall >= 5 ? 'mid' : 'low';

  const VERDICTS = {
    'hi:cheap':  ['💎', L('Quality at a reasonable price', 'جودة بسعر معقول'), 'var(--green)'],
    'hi:fair':   ['💎', L('Quality at a fair price', 'جودة بسعر عادل'), 'var(--green)'],
    'hi:rich':   ['💰', L('Great business, demanding price', 'شركة ممتازة بسعر مرتفع'), 'var(--yellow)'],
    'hi:na':     ['⭐', L('Strong fundamentals, valuation unclear', 'أساسيات قوية والتقييم غير واضح'), 'var(--accent)'],
    'mid:cheap': ['🧐', L('Average business, undemanding price', 'شركة متوسطة بسعر متواضع'), 'var(--accent)'],
    'mid:fair':  ['🔍', L('Middle of the pack', 'في منتصف المجموعة'), 'var(--text2)'],
    'mid:rich':  ['⚠️', L('Average business, rich price', 'شركة متوسطة بسعر مرتفع'), 'var(--yellow)'],
    'mid:na':    ['🔍', L('Middle of the pack', 'في منتصف المجموعة'), 'var(--text2)'],
    'low:cheap': ['⚠️', L('Cheap — but maybe for a reason', 'رخيصة — ربما لسبب'), 'var(--yellow)'],
    'low:fair':  ['⚠️', L('Weak fundamentals', 'أساسيات ضعيفة'), 'var(--red)'],
    'low:rich':  ['🚫', L('Weak fundamentals at a rich price', 'أساسيات ضعيفة بسعر مرتفع'), 'var(--red)'],
    'low:na':    ['⚠️', L('Weak fundamentals', 'أساسيات ضعيفة'), 'var(--red)'],
  };
  const [icon, label, color] = VERDICTS[`${q}:${val}`];

  // Strengths / risks from the same inputs the scores use (priority-ordered)
  const strengths = [], risks = [];
  const S = (txtEn, txtAr) => strengths.push(L(txtEn, txtAr));
  const R = (txtEn, txtAr) => risks.push(L(txtEn, txtAr));
  const yoyLast = m.yoy[m.yoy.length - 1];
  const profLast = m.prof[m.prof.length - 1] || {};
  const profPrev = m.prof[m.prof.length - 2] || {};
  const levLast = m.lev[m.lev.length - 1] || {};
  const cfLast = m.cf[m.cf.length - 1] || {};

  if (yoyLast?.revenue_growth != null) {
    if (yoyLast.revenue_growth >= 15) S(`Revenue +${yoyLast.revenue_growth.toFixed(1)}% YoY`, `نمو الإيرادات +${yoyLast.revenue_growth.toFixed(1)}% سنوياً`);
    else if (yoyLast.revenue_growth < 0) R(`Revenue shrinking (${yoyLast.revenue_growth.toFixed(1)}% YoY)`, `الإيرادات تنكمش (${yoyLast.revenue_growth.toFixed(1)}% سنوياً)`);
  }
  if (lat.net_income != null && lat.net_income <= 0) R('Loss-making (negative net income)', 'خاسرة (صافي دخل سالب)');
  if (profLast.net_margin != null) {
    if (profLast.net_margin >= 20) S(`Net margin ${profLast.net_margin.toFixed(1)}%`, `هامش صافي ${profLast.net_margin.toFixed(1)}%`);
    else if (profLast.net_margin > 0 && profLast.net_margin < 5) R(`Thin net margin (${profLast.net_margin.toFixed(1)}%)`, `هامش صافي ضعيف (${profLast.net_margin.toFixed(1)}%)`);
  }
  if (profLast.net_margin != null && profPrev.net_margin != null) {
    const d = profLast.net_margin - profPrev.net_margin;
    if (d >= 2) S('Margins expanding', 'الهوامش تتوسع');
    else if (d <= -2) R('Margins compressing', 'الهوامش تنضغط');
  }
  if (profLast.roe != null) {
    if (profLast.roe >= 20) S(`ROE ${profLast.roe.toFixed(0)}%`, `عائد على الملكية ${profLast.roe.toFixed(0)}%`);
    else if (profLast.roe >= 0 && profLast.roe < 8) R(`Low ROE (${profLast.roe.toFixed(1)}%)`, `عائد ملكية منخفض (${profLast.roe.toFixed(1)}%)`);
  }
  if (levLast.dte != null) {
    if (levLast.dte <= 0.5) S(`Low leverage (D/E ${levLast.dte.toFixed(2)}x)`, `مديونية منخفضة (${levLast.dte.toFixed(2)}x)`);
    else if (levLast.dte >= 2) R(`High leverage (D/E ${levLast.dte.toFixed(1)}x)`, `مديونية مرتفعة (${levLast.dte.toFixed(1)}x)`);
  }
  if (cfLast.fcf != null && lat.revenue > 0) {
    const fm = cfLast.fcf / lat.revenue * 100;
    if (cfLast.fcf <= 0) R('Burning cash (negative FCF)', 'تحرق النقد (تدفق حر سالب)');
    else if (fm >= 15) S(`FCF margin ${fm.toFixed(0)}%`, `هامش تدفق حر ${fm.toFixed(0)}%`);
  }
  if (cfLast.ocf_to_ni != null) {
    if (cfLast.ocf_to_ni >= 1.1) S('Earnings backed by cash (OCF > NI)', 'الأرباح مدعومة بالنقد');
    else if (cfLast.ocf_to_ni > 0 && cfLast.ocf_to_ni < 0.6) R('Weak cash conversion (OCF ≪ NI)', 'تحويل نقدي ضعيف');
  }
  if (stk.dividendYield && stk.dividendYield * 1 >= 2) S(`Dividend yield ${Number(stk.dividendYield).toFixed(1)}%`, `عائد توزيعات ${Number(stk.dividendYield).toFixed(1)}%`);
  const est = (ESTIMATES[ticker] || []).filter(e => e.epsAvg && new Date(e.date) > new Date()).sort((a, b) => a.date.localeCompare(b.date))[0];
  if (est && lat.eps_diluted > 0) {
    const g = (est.epsAvg / lat.eps_diluted - 1) * 100;
    if (g >= 15) S(`Analysts see EPS +${g.toFixed(0)}% next FY`, `المحللون يتوقعون نمو ربحية السهم +${g.toFixed(0)}%`);
    else if (g < 0) R(`Analysts see EPS declining (${g.toFixed(0)}%)`, `المحللون يتوقعون تراجع ربحية السهم (${g.toFixed(0)}%)`);
  }

  // Data confidence (completeness + freshness), not an opinion score
  let conf = 95;
  const maxYear = _dbMaxYear();
  if (m.rows.length < 4) conf -= 10;
  if (!(stk.pe > 0) && !(cfLast.fcf > 0)) conf -= 15;
  if (lat.year && maxYear && lat.year < maxYear - 1) conf -= 15;
  if (!stk.price) conf -= 10;
  conf = Math.max(30, conf);
  const confLbl = conf >= 80 ? L('High', 'عالية') : conf >= 60 ? L('Medium', 'متوسطة') : L('Low', 'منخفضة');

  return {
    icon, label, color, valLine, val, quality: q,
    strengths: strengths.slice(0, 3), risks: risks.slice(0, 3),
    conf, confLbl, ig,
    stale: !!(lat.year && maxYear && lat.year < maxYear - 1), latestYear: lat.year, maxYear,
  };
}

function renderDecisionCard(ticker, m, scores) {
  const isAr = lang === 'ar';
  const L = (en, ar) => isAr ? ar : en;
  let v;
  try { v = computeVerdict(ticker, m, scores); } catch (e) { console.warn('verdict failed', e); return ''; }
  const ig = v.ig;

  // Journal payload: rationale text saved with the pick
  _lastVerdict = {
    ticker,
    confidence: v.conf,
    summary: `[Snapshot verdict] ${v.icon} ${v.label}. ${v.valLine}. Strengths: ${v.strengths.join('; ') || '—'}. Risks: ${v.risks.join('; ') || '—'}.`
      + (ig && !ig.negative ? ` Priced-in 10y FCF growth ≈ ${ig.implied.toFixed(0)}%/yr (r=10%).` : ''),
  };

  const li = (arr, cls) => arr.length
    ? `<ul class="dc-ul ${cls}">${arr.map(x => `<li>${x}</li>`).join('')}</ul>`
    : `<div class="dc-none">${L('None flagged', 'لا شيء بارز')}</div>`;

  // "What's priced in" column
  let pricedHtml;
  if (!ig) {
    pricedHtml = `<div class="dc-none" style="margin-top:8px">${L('Not enough data (needs market cap + free cash flow).', 'لا تتوفر بيانات كافية (يتطلب القيمة السوقية والتدفق الحر).')}</div>`;
  } else if (ig.negative) {
    pricedHtml = `<div class="dc-big" style="color:var(--red)">${L('FCF negative', 'تدفق حر سالب')}</div>
      <div class="dc-sub">${L('Free cash flow is negative — today\'s price rests on a future turnaround, not current cash generation.', 'التدفق النقدي الحر سالب — السعر الحالي يراهن على تحوّل مستقبلي لا على النقد الحالي.')}</div>`;
  } else {
    const cmp = (ig.analyst != null)
      ? (ig.implied > ig.analyst + 3
          ? `<span class="dc-chip warn">${L('Price assumes MORE growth than analysts forecast', 'السعر يفترض نمواً أعلى من توقعات المحللين')}</span>`
          : ig.implied < ig.analyst - 3
            ? `<span class="dc-chip ok">${L('Price assumes LESS growth than analysts forecast', 'السعر يفترض نمواً أقل من توقعات المحللين')}</span>`
            : `<span class="dc-chip">${L('Roughly in line with analyst expectations', 'متوافق تقريباً مع توقعات المحللين')}</span>`)
      : '';
    pricedHtml = `
      <div class="dc-big">≈ ${ig.implied.toFixed(0)}%<span class="dc-big-unit">/${L('yr', 'سنة')}</span></div>
      <div class="dc-sub">${L('FCF growth for 10 years implied by the current price (10% discount, 2.5% terminal)', 'نمو التدفق الحر لعشر سنوات الذي يفترضه السعر الحالي (خصم 10%، نمو نهائي 2.5%)')}</div>
      <div class="dc-rows">
        ${ig.trailing != null ? `<div class="dc-row"><span>${L('Actual FCF growth (last 3y)', 'النمو الفعلي (آخر 3 سنوات)')}</span><b>${ig.trailing >= 0 ? '+' : ''}${ig.trailing.toFixed(0)}%/${L('yr', 'سنة')}</b></div>` : ''}
        ${ig.analyst != null ? `<div class="dc-row"><span>${L('Analysts\' fwd revenue growth', 'نمو الإيرادات المتوقع من المحللين')}</span><b>${ig.analyst >= 0 ? '+' : ''}${ig.analyst.toFixed(0)}%/${L('yr', 'سنة')}</b></div>` : ''}
        <div class="dc-row dc-row-dim"><span>${L('If discount rate 8% / 12%', 'لو معدل الخصم 8% / 12%')}</span><b>${ig.sens8.toFixed(0)}% / ${ig.sens12.toFixed(0)}%</b></div>
        <div class="dc-row dc-row-dim"><span>P/FCF</span><b>${ig.pFcf.toFixed(1)}x</b></div>
      </div>
      ${cmp}`;
  }

  return `<div class="decision-card" id="decisionCard">
    <div class="dc-col">
      <div class="dc-head">🧭 ${L('Snapshot Verdict', 'الحكم السريع')}
        <span class="dc-conf" title="${L('Based on data completeness & freshness', 'حسب اكتمال البيانات وحداثتها')}">${L('Data confidence', 'موثوقية البيانات')}: ${v.confLbl}${v.stale ? ' · ⚠ ' + L('financials lag the rest of the DB', 'البيانات متأخرة عن بقية القاعدة') : ''}</span>
      </div>
      <div class="dc-label" style="color:${v.color}">${v.icon} ${v.label}</div>
      <div class="dc-val-line">${v.valLine}</div>
      <div class="dc-lists">
        <div><div class="dc-list-h ok">✓ ${L('Strengths', 'نقاط القوة')}</div>${li(v.strengths, 'ok')}</div>
        <div><div class="dc-list-h warn">⚠ ${L('Risks', 'المخاطر')}</div>${li(v.risks, 'warn')}</div>
      </div>
    </div>
    <div class="dc-col">
      <div class="dc-head">⚖️ ${L("What's priced in?", 'ما الذي يسعّره السوق؟')}</div>
      ${pricedHtml}
    </div>
    <div class="dc-note">${L('Rule-based snapshot computed locally from the financials — educational, not investment advice.', 'لمحة آلية محسوبة محلياً من القوائم المالية — لأغراض تعليمية وليست نصيحة استثمارية.')}</div>
  </div>`;
}

function loadTickerIntoDash(ticker) {
  _dashTarget = document.getElementById('dashPanelContent');
  loadTicker(ticker);
}

function loadTicker(ticker) {
  if (!FILE_LOADED || !TICKERS.length) return;

  // If dashboard panel is open, redirect into it instead of replacing mainArea
  const dashPanel = document.getElementById('dashPanel');
  if (!_dashTarget && dashPanel && !dashPanel.classList.contains('hidden')) {
    loadTickerIntoDash(ticker);
    activeTicker = ticker;
    renderSidebarTickers();
    return;
  }

  activeTicker = ticker;
  renderSidebarTickers();

  const target = _dashTarget || getMainTarget();

  const rows = ANNUAL[ticker];
  if (!rows || !rows.length) {
    target.innerHTML = `<div class="welcome"><div class="welcome-eyebrow">⚠️</div><h1 style="font-size:28px">No data for ${ticker}</h1></div>`;
    return;
  }

  target.innerHTML = `<div class="loader" id="dbLoader"><div class="spin"></div><span class="loader-txt">${T('loaderTxt')}</span></div>`;

  setTimeout(() => {
    const m = calcMetrics(rows);
    const scores = calcScores(m);
    const insights = genInsights(ticker, m, scores);
    const stk = STOCK[ticker] || {};
    const qData = QUARTERLY[ticker] || [];
    currentDash = {ticker, rows, m, scores};

    const lat = m.latest;
    const inCmp = cmpList.includes(ticker);

    // Score cards
    const scoreCats = [
      {k:'growth',lbl:T('growth')},{k:'profitability',lbl:T('profitability')},
      {k:'health',lbl:T('health')},{k:'cashflow',lbl:T('cashflow')},{k:'overall',lbl:T('overallScore'),cls:' sc-overall'}
    ];
    const scHtml = scoreCats.map(s=>{
      const v=scores[s.k], col=sCol(v);
      return `<div class="sc${s.cls||''}" style="--c:${col}">
        <div class="sc-lbl">${s.k==='overall'?'⭐ ':''}${s.lbl}</div>
        <div class="sc-num">${v}<span class="sc-den">/10</span></div>
        <div class="sc-bar"><div class="sc-fill" style="width:${v*10}%"></div></div>
      </div>`;
    }).join('');

    // KPIs
    const kpiData = [
      {k:'revenue',v:fM(lat.revenue)},{k:'netIncome',v:fM(lat.net_income)},
      {k:'totalAssets',v:fM(lat.total_assets)},{k:'equity',v:fM(lat.equity)},
      {k:'ocf',v:fM(lat.operating_cash_flow)},{k:'fcf',v:fM(lat.free_cash_flow)},
      {k:'ebitda',v:fM(lat.ebitda)},{k:'eps',v:lat.eps_diluted!=null?'$'+lat.eps_diluted.toFixed(2):T('noData')},
    ];
    const kpiHtml = kpiData.map(k=>`<div class="kpi">
      <div class="kpi-lbl">${T(k.k)}</div>
      <div class="kpi-val" style="color:${k.v.startsWith('(')?'var(--red)':'var(--text)'}">${k.v}</div>
      <div class="kpi-sub">${T('latestYr')}</div>
    </div>`).join('');

    // Stock bar
    const recCls = {buy:'rec-buy',hold:'rec-hold',sell:'rec-sell'}[stk.recommendation] || '';
    const stkHtml = (stk.price||stk.marketCap) ? `<div class="stock-bar" id="stockBar">
      ${stk.price?`<div class="sb-item"><span class="sb-lbl">${T('price')}</span><span class="sb-val" id="livePrice">$${stk.price.toFixed(2)}</span></div><div class="sb-divider"></div>`:'<div class="sb-item"><span class="sb-lbl">${T("price")}</span><span class="sb-val" id="livePrice">—</span></div><div class="sb-divider"></div>'}
      <div class="sb-item" id="liveChangeWrap"><span class="sb-change" id="liveChange">—</span></div><div class="sb-divider"></div>
      ${stk.marketCap?`<div class="sb-item"><span class="sb-lbl">${T('mktCap')}</span><span class="sb-val" id="liveMcap">${fM(stk.marketCap)}</span></div><div class="sb-divider"></div>`:`<div class="sb-item"><span class="sb-lbl">${T('mktCap')}</span><span class="sb-val" id="liveMcap">—</span></div><div class="sb-divider"></div>`}
      ${stk.pe?`<div class="sb-item"><span class="sb-lbl">${T('pe')}</span><span class="sb-val" id="livePe">${stk.pe.toFixed(1)}x</span></div><div class="sb-divider"></div>`:`<div class="sb-item"><span class="sb-lbl">${T('pe')}</span><span class="sb-val" id="livePe">—</span></div><div class="sb-divider"></div>`}
      ${stk.currentRatio?`<div class="sb-item"><span class="sb-lbl">${T('currentRatio')}</span><span class="sb-val">${stk.currentRatio.toFixed(2)}</span></div><div class="sb-divider"></div>`:''}
      ${stk.dividendYield?`<div class="sb-item"><span class="sb-lbl">Yield</span><span class="sb-val">${stk.dividendYield.toFixed(2)}%</span></div><div class="sb-divider"></div>`:''}
      ${stk.recommendation?`<div class="sb-item"><span class="sb-lbl">${T('recommendation')}</span><span class="sb-val ${recCls}" style="text-transform:capitalize">${stk.recommendation}</span></div>`:''}
      <div class="sb-live hidden" id="liveTag"><div class="sb-live-dot"></div> LIVE</div>
      <button class="sb-journal-btn" onclick="quickSaveCurrentTicker()" title="${lang==='ar'?'احفظ في السجل':'Save to Journal'}">📓 ${lang==='ar'?'احفظ':'Save'}</button>
    </div>` : `<div class="stock-bar" id="stockBar">
      <div class="sb-item"><span class="sb-lbl">${T('price')}</span><span class="sb-val" id="livePrice">—</span></div><div class="sb-divider"></div>
      <div class="sb-item"><span class="sb-change" id="liveChange">—</span></div><div class="sb-divider"></div>
      <div class="sb-item"><span class="sb-lbl">${T('mktCap')}</span><span class="sb-val" id="liveMcap">—</span></div>
      <div class="sb-live hidden" id="liveTag"><div class="sb-live-dot"></div> LIVE</div>
      <button class="sb-journal-btn" onclick="quickSaveCurrentTicker()" title="${lang==='ar'?'احفظ في السجل':'Save to Journal'}">📓 ${lang==='ar'?'احفظ':'Save'}</button>
    </div>`;

    // 52-week range bar
    const rangeHtml = (stk.fiftyTwoWeekLow && stk.fiftyTwoWeekHigh && stk.price) ? (() => {
      const pct = ((stk.price - stk.fiftyTwoWeekLow) / (stk.fiftyTwoWeekHigh - stk.fiftyTwoWeekLow) * 100).toFixed(0);
      return `<div id="liveRangeWrap" style="background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:16px;flex-wrap:wrap">
        <span style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;white-space:nowrap">52W Range</span>
        <div id="liveRangeBar" style="display:contents">
        <span style="font-size:13px;color:var(--text2)">$${stk.fiftyTwoWeekLow.toFixed(2)}</span>
        <div style="flex:1;min-width:80px;height:6px;background:var(--border);border-radius:3px;position:relative">
          <div style="position:absolute;left:${pct}%;top:-3px;width:12px;height:12px;background:var(--accent);border-radius:50%;transform:translateX(-50%);box-shadow:0 0 6px rgba(59,130,246,.5)"></div>
          <div style="height:100%;width:${pct}%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:3px"></div>
        </div>
        <span style="font-size:13px;color:var(--text2)">$${stk.fiftyTwoWeekHigh.toFixed(2)}</span>
        <span style="font-size:12px;color:var(--text3)">${pct}%</span>
        </div>
      </div>`;
    })() : `<div id="liveRangeWrap" style="background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:16px;flex-wrap:wrap">
      <span style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.08em">52W Range</span>
      <div id="liveRangeBar" style="display:contents"><span style="font-size:13px;color:var(--text3)">Loading...</span></div>
    </div>`;

    // YoY table
    const yoyHtml = m.yoy.map(y=>`<tr>
      <td>${y.year}</td>
      <td class="${pc(y.revenue_growth)}">${fP(y.revenue_growth)}</td>
      <td class="${pc(y.ni_growth)}">${fP(y.ni_growth)}</td>
      <td class="${pc(y.ocf_growth)}">${fP(y.ocf_growth)}</td>
      <td class="${pc(y.eps_growth)}">${fP(y.eps_growth)}</td>
    </tr>`).join('');

    // Profitability table
    const profHtml = m.prof.map((p,i)=>`<tr>
      <td>${p.year}</td>
      <td class="${pc(p.net_margin)}">${p.net_margin!=null?p.net_margin.toFixed(1)+'%':'—'}</td>
      <td class="${pc(p.op_margin)}">${p.op_margin!=null?p.op_margin.toFixed(1)+'%':'—'}</td>
      <td class="${pc(p.ebitda_margin)}">${p.ebitda_margin!=null?p.ebitda_margin.toFixed(1)+'%':'—'}</td>
      <td class="${pc(p.roe)}">${p.roe!=null?p.roe.toFixed(1)+'%':'—'}</td>
      <td class="${pc(p.roa)}">${p.roa!=null?p.roa.toFixed(1)+'%':'—'}</td>
      <td>${m.lev[i]?.dte!=null?fR(m.lev[i].dte):'—'}</td>
    </tr>`).join('');

    // Quarterly table
    const qHtml = qData.length ? `<div class="tbl-sec">
      <div class="sec-hdr"><div class="sec-title">📋 ${T('qTable')}</div></div>
      <div class="tbl-card"><div class="tbl-scroll"><table>
        <thead><tr><th>${T('date')}</th><th>${T('revenue')}</th><th>${T('netIncome')}</th><th>${T('opIncome')}</th><th>${T('eps')}</th><th>${T('ocf')}</th></tr></thead>
        <tbody>${qData.map(q=>`<tr>
          <td>${q.date}</td><td>${fM(q.revenue)}</td>
          <td class="${pc(q.net_income)}">${fM(q.net_income)}</td>
          <td class="${pc(q.operating_income)}">${fM(q.operating_income)}</td>
          <td class="${pc(q.eps_diluted)}">${q.eps_diluted!=null?'$'+q.eps_diluted.toFixed(2):'—'}</td>
          <td class="${pc(q.operating_cash_flow)}">${fM(q.operating_cash_flow)}</td>
        </tr>`).join('')}</tbody>
      </table></div></div>
    </div>` : '';

    // Insights
    const insHtml = (insights[lang]||insights.en).map(i=>`<div class="ins-item ${i.type}">
      <span class="ins-icon">${i.icon}</span><span class="ins-txt">${i.text}</span>
    </div>`).join('');

    // Decision card (rule-based verdict + reverse-DCF; token-free)
    const decisionHtml = renderDecisionCard(ticker, m, scores);

    (_dashTarget || document.getElementById('dashFullView') || document.getElementById('mainArea')).innerHTML = `
    <div class="db">
      <div class="db-hdr">
        <div class="db-left">
          <div class="db-ticker">${ticker}</div>
          <div class="db-meta">
            <div class="db-score-line">${T('overallScore')}: <span style="color:${sCol(scores.overall)};font-weight:700">${scores.overall}/10</span></div>
            <div class="db-info">${rows.length} ${T('yearsData')} · ${T('latestYr')}: ${lat.year||'—'}</div>
            ${stk.sector?`<div class="db-sector"><span class="tag tag-blue">${stk.sector}</span>${stk.industry?` <span style="font-size:11px;color:var(--text3)">${stk.industry}</span>`:''}</div>`:''}
          </div>
        </div>
        <div class="db-actions">
          <button class="act-btn" onclick="backToChat()" style="color:var(--accent);border-color:var(--accent)">💬 ${lang==='ar'?'المحادثة':'Chat'}</button>
          <button class="act-btn" onclick="toggleWatchlist('${ticker}')" id="wlBtn" style="${watchlist.includes(ticker)?'color:var(--accent2);border-color:var(--accent2)':''}">
            ${watchlist.includes(ticker)?(lang==='ar'?'👁️ في المراقبة':'👁️ Watching'):(lang==='ar'?'👁️ راقب':'👁️ Watch')}
          </button>
          <button class="act-btn add-cmp-btn${inCmp?' added':''}" onclick="toggleCmp('${ticker}')" id="addCmpBtn">
            ${inCmp?T('addedCmp'):T('addCmp')}
          </button>
          <button class="act-btn" onclick="doExportCSV('${ticker}')">📥 ${T('exportCSV')}</button>
          <button class="act-btn" onclick="window.print()">🖨️ ${T('exportPDF')}</button>
        </div>
      </div>

      ${stkHtml}
      ${rangeHtml}
      ${decisionHtml}

      <div class="price-chart-wrap" id="priceChartWrap">
        <div class="price-chart-header">
          <div>
            <span class="price-chart-title">${lang==='ar'?'سجل الأسعار':'Price History'}</span>
            <span id="priceChartChange" style="font-size:12px;font-weight:600;margin-left:8px">—</span>
          </div>
          <div class="price-chart-tabs">
            <button class="price-tab" data-range="1d" onclick="loadPriceChart('${ticker}','1d')">1D</button>
            <button class="price-tab" data-range="5d" onclick="loadPriceChart('${ticker}','5d')">5D</button>
            <button class="price-tab active" data-range="1m" onclick="loadPriceChart('${ticker}','1m')">1M</button>
            <button class="price-tab" data-range="6m" onclick="loadPriceChart('${ticker}','6m')">6M</button>
            <button class="price-tab" data-range="1y" onclick="loadPriceChart('${ticker}','1y')">1Y</button>
            <button class="price-tab" data-range="5y" onclick="loadPriceChart('${ticker}','5y')">5Y</button>
          </div>
        </div>
        <div style="height:200px;position:relative"><canvas id="priceHistoryChart" onmouseleave="if(charts['priceHistory']){charts['priceHistory']._crosshairX=null;charts['priceHistory']._crosshairY=null;charts['priceHistory'].draw()}"></canvas></div>
      </div>

      <div class="score-row">${scHtml}</div>
      <div class="kpi-grid">${kpiHtml}</div>

      <div class="ai-block" id="agentPanel">
        <div class="agent-hdr">
          <div class="agent-hdr-left">
            <span class="ai-tag" style="margin:0">${T('aiTag')}</span>
            <span class="ai-model-badge">${getModelMeta(DEEPSEEK_MODEL).short}</span>
          </div>
          <div class="agent-hdr-actions">
            <select id="aiLengthSelect" style="padding:4px 8px;border-radius:8px;border:1px solid var(--border);background:var(--surface2);color:var(--text2);font-size:10px;font-family:var(--font);cursor:pointer" title="Analysis length">
              <option value="brief">${lang==='ar'?'مختصر':'Brief'}</option>
              <option value="standard" selected>${lang==='ar'?'عادي':'Standard'}</option>
              <option value="deep">${lang==='ar'?'معمّق':'Deep'}</option>
            </select>
            <button class="ai-generate" id="aiGenBtn" onclick="agentStart('${ticker}')">${isAiReady() ? (lang==='ar'?'✦ تحليل ذكي':'✦ Analyse') : (lang==='ar'?'🔐 سجّل دخولك':'🔐 Sign in for AI')}</button>
            <button class="ai-regen hidden" id="aiRegenBtn" onclick="agentStart('${ticker}')">${T('aiRegen')}</button>
          </div>
        </div>
        <div class="agent-conv" id="agentConv" data-empty="${lang==='ar'?'اضغط «✦ تحليل ذكي» لتفعيل الوكيل الذكي.\\nسيقوم بتحليل البيانات واكتشاف المخاطر ومقارنة الأقران تلقائياً.':'Click ✦ Analyse to activate the AI agent.\\nIt will examine the data, spot risks, compare peers, and suggest next steps.'}">
        </div>
        <div class="agent-actions hidden" id="agentActions"></div>
        <div class="agent-chat-row">
          <input class="agent-chat-input" id="agentChatInput" placeholder="${lang==='ar'?'اسأل عن هذه الشركة...':'Ask about this company...'}" onkeydown="if(event.key==='Enter')agentChat('${ticker}')"/>
          <button class="agent-chat-send" id="agentChatSend" onclick="agentChat('${ticker}')">${lang==='ar'?'إرسال':'Send'}</button>
        </div>
        <div class="agent-status" id="agentStatus">
          <div class="agent-status-dot idle"></div>
          <span id="agentStatusTxt">${lang==='ar'?'الوكيل في وضع الاستعداد':'Agent standby'}</span>
        </div>
      </div>

      ${(() => {
        const est = ESTIMATES[ticker];
        if (!est || !est.length) return '';
        // Find the nearest future estimate
        const now = new Date();
        const future = est.filter(e => new Date(e.date) > now).sort((a,b) => a.date.localeCompare(b.date));
        const next = future[0] || est[est.length - 1];
        if (!next) return '';
        const isAr = lang === 'ar';
        const latestRev = m.latest.revenue;
        const latestEPS = m.latest.eps_diluted;
        const revVs = latestRev && next.revenueAvg ? ((latestRev / next.revenueAvg - 1) * 100) : null;
        const epsVs = latestEPS && next.epsAvg ? ((latestEPS / next.epsAvg - 1) * 100) : null;
        return `<div style="margin-bottom:18px">
          <div class="sec-hdr"><div class="sec-title">🎯 ${isAr ? 'توقعات المحللين' : 'Wall Street Estimates'} <span style="font-size:11px;color:var(--text3);font-weight:400">(${next.date?.slice(0,4) || ''})</span></div></div>
          <div class="est-grid">
            <div class="est-card">
              <div class="est-label">${isAr ? 'الإيرادات المتوقعة' : 'Est. Revenue'}</div>
              <div class="est-value">${next.revenueAvg ? fM(next.revenueAvg) : '—'}</div>
              ${next.revenueLow && next.revenueHigh ? `<div class="est-range">${fM(next.revenueLow)} – ${fM(next.revenueHigh)}</div>` : ''}
              ${next.numAnalystsRevenue ? `<div class="est-analysts">${next.numAnalystsRevenue} analysts</div>` : ''}
            </div>
            <div class="est-card">
              <div class="est-label">${isAr ? 'ربحية السهم المتوقعة' : 'Est. EPS'}</div>
              <div class="est-value">$${next.epsAvg?.toFixed(2) || '—'}</div>
              ${next.epsLow && next.epsHigh ? `<div class="est-range">$${next.epsLow?.toFixed(2)} – $${next.epsHigh?.toFixed(2)}</div>` : ''}
              ${next.numAnalystsEps ? `<div class="est-analysts">${next.numAnalystsEps} analysts</div>` : ''}
            </div>
            <div class="est-card">
              <div class="est-label">${isAr ? 'صافي الدخل المتوقع' : 'Est. Net Income'}</div>
              <div class="est-value">${next.netIncomeAvg ? fM(next.netIncomeAvg) : '—'}</div>
            </div>
          </div>
        </div>`;
      })()}

      <div style="margin-bottom:18px">
        <div class="sec-hdr"><div class="sec-title">📰 ${lang==='ar' ? 'آخر الأخبار' : 'Latest News'}</div></div>
        <div class="news-feed" id="newsFeed"><div class="news-loading">${lang==='ar'?'جار تحميل الأخبار...':'Loading news...'}</div></div>
      </div>

      <div style="margin-bottom:18px">
        <div class="sec-hdr"><div class="sec-title">🔗 ${lang==='ar' ? 'الشركات المرتبطة' : 'Related Companies'}</div></div>
        <div class="kg-panel" id="kgPanel"><div class="news-loading">${lang==='ar'?'جار تحميل الرسم البياني...':'Loading knowledge graph...'}</div></div>
      </div>

      <div class="charts-2">
        <div class="chart-card"><div class="chart-title">📈 ${T('revChart')}</div><div class="chart-wrap"><canvas id="cRev"></canvas></div></div>
        <div class="chart-card"><div class="chart-title">💵 ${T('cfChart')}</div><div class="chart-wrap"><canvas id="cCF"></canvas></div></div>
        <div class="chart-card"><div class="chart-title">📊 ${T('yoyChart')}</div><div class="chart-wrap"><canvas id="cYoY"></canvas></div></div>
        <div class="chart-card"><div class="chart-title">🎯 ${T('radarChart')}</div><div class="chart-wrap"><canvas id="cRadar"></canvas></div></div>
      </div>

      <div class="tbl-sec">
        <div class="sec-hdr"><div class="sec-title">📅 ${T('yoyTable')}</div></div>
        <div class="tbl-card"><div class="tbl-scroll"><table>
          <thead><tr><th>${T('yr')}</th><th>${T('revGrowth')}</th><th>${T('niGrowth')}</th><th>${T('ocfGrowth')}</th><th>EPS Growth</th></tr></thead>
          <tbody>${yoyHtml||'<tr><td colspan="5" class="neu" style="text-align:center">—</td></tr>'}</tbody>
        </table></div></div>
      </div>

      <div class="tbl-sec">
        <div class="sec-hdr"><div class="sec-title">💰 ${T('profTable')}</div></div>
        <div class="tbl-card"><div class="tbl-scroll"><table>
          <thead><tr><th>${T('yr')}</th><th>${T('netMargin')}</th><th>${T('opMargin')}</th><th>EBITDA Margin</th><th>${T('roe')}</th><th>${T('roa')}</th><th>${T('dte')}</th></tr></thead>
          <tbody>${profHtml||'<tr><td colspan="7" class="neu" style="text-align:center">—</td></tr>'}</tbody>
        </table></div></div>
      </div>

      ${qHtml}

      <div class="sec-hdr"><div class="sec-title">🔍 Insights</div></div>
      <div class="ins-grid">${insHtml}</div>
    </div>`;

    setTimeout(() => {
      drawCharts(ticker, rows, m);
      // Fetch live data in background (non-blocking)
      if (FMP_QUOTE_KEY) {
        fetchLiveQuote(ticker).then(q => updateStockBarWithLive(ticker, q));
        fetchLiveNews(ticker).then(async news => {
          if (!news?.length) { renderNewsSection(news); return; }
          // 1) Render news IMMEDIATELY so it always shows, even if sentiment scoring
          //    hangs or fails. (Sentiment shows as 'unknown' until scored.)
          renderNewsSection(news);
          // 2) Enhance with AI sentiment in the background — never let a failure
          //    here block the news from displaying.
          try {
            const scored = await analyseNewsSentiment(news, ticker);
            renderNewsSection(scored);
            if (isAiReady() && agentMemory[ticker]) {
              agentMemory[ticker].sentiment = scored.reduce((s, n) => s + (n.sentimentScore || 0), 0);
            }
          } catch (e) {
            console.warn('News sentiment scoring failed (news still shown):', e);
          }
        }).catch(e => {
          console.warn('News fetch failed:', e);
          renderNewsSection([]);
        });
        renderKgPanel(ticker);
        loadPriceChart(ticker, '1m');
      }
    }, 60);
    _dashTarget = null;
  }, 150);
}

// ── SECTOR DASHBOARD ─────────────────────────────────────────────────────────
function computeSectorData(sector) {
  const tickers = TICKERS.filter(t => STOCK[t]?.sector === sector && ANNUAL[t]?.length);
  const companies = {};
  tickers.forEach(t => {
    const m = calcMetrics(ANNUAL[t]);
    const scores = calcScores(m);
    companies[t] = { m, scores, stk: STOCK[t] || {}, rows: ANNUAL[t] };
  });

  // Aggregate scores
  const avgOf = (key) => {
    const vals = tickers.map(t => companies[t].scores[key]).filter(v => v != null);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  const avgScores = {
    growth: avgOf('growth'), profitability: avgOf('profitability'),
    health: avgOf('health'), cashflow: avgOf('cashflow'), overall: avgOf('overall')
  };

  // Aggregate metrics
  const allProf = tickers.flatMap(t => companies[t].m.prof);
  const allLev = tickers.flatMap(t => companies[t].m.lev);
  const allYoy = tickers.flatMap(t => companies[t].m.yoy);
  const avgMetric = (arr, key) => {
    const vals = arr.map(r => r[key]).filter(v => v != null && !isNaN(v));
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  const avgNetMargin = avgMetric(allProf, 'net_margin');
  const avgROE = avgMetric(allProf, 'roe');
  const avgDTE = avgMetric(allLev, 'dte');
  const avgRevGrowth = avgMetric(allYoy, 'revenue_growth');

  // Industry breakdown
  const industries = {};
  tickers.forEach(t => {
    const ind = STOCK[t]?.industry || 'Other';
    if (!industries[ind]) industries[ind] = [];
    industries[ind].push(t);
  });

  // Sort by overall score
  const sorted = [...tickers].sort((a, b) => (companies[b].scores.overall || 0) - (companies[a].scores.overall || 0));
  const top5 = sorted.slice(0, 5);
  const bottom5 = sorted.slice(-5).reverse();

  // Total market cap
  const totalMCap = tickers.reduce((sum, t) => sum + (STOCK[t]?.marketCap || 0), 0);

  return { tickers, companies, avgScores, avgNetMargin, avgROE, avgDTE, avgRevGrowth, industries, sorted, top5, bottom5, totalMCap };
}

function renderSectorDashboard(sector) {
  getMainTarget().innerHTML = `<div class="loader"><div class="spin"></div><span class="loader-txt">${T('loaderTxt')}</span></div>`;

  setTimeout(() => {
    const d = computeSectorData(sector);
    if (!d.tickers.length) {
      getMainTarget().innerHTML = `<div class="welcome"><h1 style="font-size:28px">No data for ${sector}</h1></div>`;
      return;
    }

    const isAr = lang === 'ar';

    // Aggregate stats row
    const statsHtml = [
      { lbl: isAr ? 'الشركات' : 'Companies', val: d.tickers.length, sub: isAr ? 'في هذا القطاع' : 'in sector' },
      { lbl: isAr ? 'الدرجة الإجمالية' : 'Avg Score', val: d.avgScores.overall?.toFixed(1) || '—', sub: '/10', color: sCol(d.avgScores.overall || 5) },
      { lbl: isAr ? 'نمو الإيرادات' : 'Avg Rev Growth', val: d.avgRevGrowth != null ? (d.avgRevGrowth > 0 ? '+' : '') + d.avgRevGrowth.toFixed(1) + '%' : '—', sub: isAr ? 'سنوي' : 'annual', color: d.avgRevGrowth > 0 ? 'var(--green)' : 'var(--red)' },
      { lbl: isAr ? 'هامش صافي' : 'Avg Net Margin', val: d.avgNetMargin != null ? d.avgNetMargin.toFixed(1) + '%' : '—', sub: isAr ? 'متوسط القطاع' : 'sector avg', color: d.avgNetMargin > 10 ? 'var(--green)' : d.avgNetMargin > 0 ? 'var(--accent)' : 'var(--red)' },
      { lbl: isAr ? 'الدين/الملكية' : 'Avg D/E', val: d.avgDTE != null ? d.avgDTE.toFixed(2) + 'x' : '—', sub: isAr ? 'الرافعة المالية' : 'leverage', color: d.avgDTE < 1 ? 'var(--green)' : d.avgDTE < 2 ? 'var(--accent)' : 'var(--red)' },
    ].map(s => `<div class="sec-stat">
      <div class="sec-stat-lbl">${s.lbl}</div>
      <div class="sec-stat-val" ${s.color ? `style="color:${s.color}"` : ''}>${s.val}</div>
      <div class="sec-stat-sub">${s.sub}</div>
    </div>`).join('');

    // Score breakdown row
    const scoreBreakdown = ['growth', 'profitability', 'health', 'cashflow'].map(k => {
      const v = d.avgScores[k];
      const col = sCol(v || 5);
      const lbl = T(k);
      return `<div class="sc" style="--c:${col}">
        <div class="sc-lbl">${lbl}</div>
        <div class="sc-num">${v?.toFixed(1) || '—'}<span class="sc-den">/10</span></div>
        <div class="sc-bar"><div class="sc-fill" style="width:${(v || 0) * 10}%"></div></div>
      </div>`;
    }).join('');

    // Industry pills
    const indSorted = Object.entries(d.industries).sort((a, b) => b[1].length - a[1].length);
    const indPillsHtml = indSorted.map(([ind, tks]) =>
      `<div class="ind-pill" onclick="filterSectorByIndustry('${ind.replace(/'/g, "\\'")}','${sector.replace(/'/g, "\\'")}')">${ind} <span class="ind-pill-cnt">${tks.length}</span></div>`
    ).join('');

    // Top 5 performers
    const topHtml = d.top5.map((t, i) => {
      const s = d.companies[t].scores.overall;
      return `<div class="perf-row" onclick="loadTicker('${t}')">
        <div class="perf-row-left">
          <span class="perf-rank">${i + 1}</span>
          <div><div class="perf-ticker">${t}</div><div class="perf-industry">${STOCK[t]?.industry || ''}</div></div>
        </div>
        <div class="perf-score" style="color:${sCol(s)}">${s}/10</div>
      </div>`;
    }).join('');

    // Bottom 5
    const botHtml = d.bottom5.map((t, i) => {
      const s = d.companies[t].scores.overall;
      return `<div class="perf-row" onclick="loadTicker('${t}')">
        <div class="perf-row-left">
          <span class="perf-rank">${d.tickers.length - i}</span>
          <div><div class="perf-ticker">${t}</div><div class="perf-industry">${STOCK[t]?.industry || ''}</div></div>
        </div>
        <div class="perf-score" style="color:${sCol(s)}">${s}/10</div>
      </div>`;
    }).join('');

    // Heatmap cells
    const hmHtml = d.sorted.map(t => {
      const s = d.companies[t].scores.overall;
      const col = sCol(s);
      const alpha = Math.max(0.15, Math.min(0.5, s / 15));
      return `<div class="hm-cell" onclick="loadTicker('${t}')" style="background:${col}${Math.round(alpha * 255).toString(16).padStart(2, '0')}">
        <div class="hm-cell-ticker">${t}</div>
        <div class="hm-cell-score" style="color:${col}">${s}</div>
      </div>`;
    }).join('');

    // Industry table
    const indTableHtml = indSorted.map(([ind, tks]) => {
      const indScores = tks.map(t => d.companies[t].scores.overall).filter(v => v != null);
      const indAvg = indScores.length ? (indScores.reduce((a, b) => a + b, 0) / indScores.length) : null;
      const indMargins = tks.flatMap(t => d.companies[t].m.prof.map(p => p.net_margin)).filter(v => v != null);
      const indAvgMargin = indMargins.length ? (indMargins.reduce((a, b) => a + b, 0) / indMargins.length) : null;
      const indGrowths = tks.flatMap(t => d.companies[t].m.yoy.map(y => y.revenue_growth)).filter(v => v != null);
      const indAvgGrowth = indGrowths.length ? (indGrowths.reduce((a, b) => a + b, 0) / indGrowths.length) : null;
      const best = tks.reduce((a, b) => (d.companies[a]?.scores.overall || 0) >= (d.companies[b]?.scores.overall || 0) ? a : b);
      return `<tr>
        <td style="font-weight:600">${ind}</td>
        <td style="text-align:center">${tks.length}</td>
        <td style="text-align:center;color:${sCol(indAvg || 5)};font-weight:700">${indAvg?.toFixed(1) || '—'}</td>
        <td style="text-align:center" class="${pc(indAvgGrowth)}">${indAvgGrowth != null ? (indAvgGrowth > 0 ? '+' : '') + indAvgGrowth.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center" class="${pc(indAvgMargin)}">${indAvgMargin != null ? indAvgMargin.toFixed(1) + '%' : '—'}</td>
        <td style="text-align:center;cursor:pointer;color:var(--accent)" onclick="loadTicker('${best}')">${best}</td>
      </tr>`;
    }).join('');

    getMainTarget().innerHTML = `
    <div class="sec-dash">
      <div class="sec-dash-hdr">
        <div class="sec-dash-title">${sector}</div>
        <div class="sec-dash-sub">${d.tickers.length} ${isAr ? 'شركة' : 'companies'} · ${indSorted.length} ${isAr ? 'صناعة' : 'industries'}${d.totalMCap ? ' · ' + fM(d.totalMCap) + (isAr ? ' القيمة السوقية الإجمالية' : ' total market cap') : ''}</div>
      </div>

      <div class="sec-stats-row">${statsHtml}</div>
      <div class="score-row">${scoreBreakdown}</div>

      <div class="sec-hdr" style="margin-top:18px"><div class="sec-title">${isAr ? '🏭 الصناعات' : '🏭 Industries'}</div></div>
      <div class="ind-pills">${indPillsHtml}</div>

      <div class="perf-grid">
        <div class="perf-card">
          <div class="perf-card-title top">${isAr ? '🏆 الأفضل أداءً' : '🏆 Top performers'}</div>
          ${topHtml}
        </div>
        <div class="perf-card">
          <div class="perf-card-title bottom">${isAr ? '⚠️ الأضعف أداءً' : '⚠️ Underperformers'}</div>
          ${botHtml}
        </div>
      </div>

      <div class="tbl-sec">
        <div class="sec-hdr"><div class="sec-title">${isAr ? '📊 تحليل الصناعات' : '📊 Industry breakdown'}</div></div>
        <div class="tbl-card"><div class="tbl-scroll"><table>
          <thead><tr>
            <th>${isAr ? 'الصناعة' : 'Industry'}</th>
            <th style="text-align:center">#</th>
            <th style="text-align:center">${isAr ? 'الدرجة' : 'Avg Score'}</th>
            <th style="text-align:center">${isAr ? 'النمو' : 'Avg Growth'}</th>
            <th style="text-align:center">${isAr ? 'الهامش' : 'Avg Margin'}</th>
            <th style="text-align:center">${isAr ? 'الأفضل' : 'Top Pick'}</th>
          </tr></thead>
          <tbody>${indTableHtml}</tbody>
        </table></div></div>
      </div>

      <div class="sec-hdr"><div class="sec-title">${isAr ? '🗺️ خريطة الأداء' : '🗺️ Performance heatmap'}</div></div>
      <div class="sec-heatmap">${hmHtml}</div>

      <div class="charts-2">
        <div class="chart-card"><div class="chart-title">${isAr ? '📈 توزيع الدرجات' : '📈 Score distribution'}</div><div class="chart-wrap"><canvas id="cSecScoreDist"></canvas></div></div>
        <div class="chart-card"><div class="chart-title">${isAr ? '🎯 الربحية مقابل النمو' : '🎯 Profitability vs Growth'}</div><div class="chart-wrap"><canvas id="cSecScatter"></canvas></div></div>
      </div>
    </div>`;

    setTimeout(() => {
      destroyCharts();
      Chart.defaults.color = getChartColors().legend;
      Chart.defaults.borderColor = getChartColors().grid;

      // Score distribution histogram
      const bins = [0, 2, 4, 5, 6, 7, 8, 10];
      const binLabels = bins.slice(0, -1).map((b, i) => `${b}-${bins[i + 1]}`);
      const binCounts = binLabels.map((_, i) => d.sorted.filter(t => {
        const s = d.companies[t].scores.overall;
        return s >= bins[i] && s < bins[i + 1];
      }).length);
      // Last bin includes 10
      binCounts[binCounts.length - 1] += d.sorted.filter(t => d.companies[t].scores.overall === 10).length;
      const binColors = binLabels.map((_, i) => {
        const mid = (bins[i] + bins[i + 1]) / 2;
        return sCol(mid) + 'cc';
      });

      mk('cSecScoreDist', {
        type: 'bar',
        data: {
          labels: binLabels,
          datasets: [{ data: binCounts, backgroundColor: binColors, borderColor: binColors.map(c => c.replace('cc', '')), borderWidth: 1, borderRadius: 6 }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { ...scOpts(), y: { ...scOpts().y, title: { display: true, text: isAr ? 'عدد الشركات' : 'Companies', color: getChartColors().legend } } }, plugins: { legend: { display: false } } }
      });

      // Profitability vs Growth scatter
      const scatterData = d.tickers.map(t => {
        const c = d.companies[t];
        const lastProf = c.m.prof[c.m.prof.length - 1];
        const lastYoy = c.m.yoy[c.m.yoy.length - 1];
        return { x: lastYoy?.revenue_growth ?? null, y: lastProf?.net_margin ?? null, ticker: t };
      }).filter(p => p.x != null && p.y != null);

      mk('cSecScatter', {
        type: 'scatter',
        data: {
          datasets: [{
            data: scatterData.map(p => ({ x: p.x, y: p.y })),
            backgroundColor: scatterData.map(p => sCol(d.companies[p.ticker].scores.overall) + '99'),
            borderColor: scatterData.map(p => sCol(d.companies[p.ticker].scores.overall)),
            borderWidth: 1, pointRadius: 6, pointHoverRadius: 9
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          scales: {
            x: { ...scOpts().x, title: { display: true, text: isAr ? 'نمو الإيرادات %' : 'Revenue Growth %', color: getChartColors().legend } },
            y: { ...scOpts().y, title: { display: true, text: isAr ? 'هامش صافي %' : 'Net Margin %', color: getChartColors().legend } }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) => {
                  const p = scatterData[ctx.dataIndex];
                  return `${p.ticker}: Growth ${p.x.toFixed(1)}%, Margin ${p.y.toFixed(1)}%`;
                }
              }
            }
          }
        }
      });
    }, 60);
  }, 150);
}

function filterSectorByIndustry(industry, sector) {
  renderIndustryDashboard(industry, sector);
}

// ── COMPARE ───────────────────────────────────────────────────────────────────
function toggleCmp(ticker) {
  if (cmpList.includes(ticker)) {
    cmpList = cmpList.filter(t => t !== ticker);
  } else if (cmpList.length < 5) {
    cmpList.push(ticker);
  }
  updateCmpBar();
  // Update button on current dashboard
  const btn = document.getElementById('addCmpBtn');
  if (btn) {
    const inCmp = cmpList.includes(ticker);
    btn.textContent = inCmp ? T('addedCmp') : T('addCmp');
    btn.classList.toggle('added', inCmp);
  }
}

function clearCmp() { cmpList = []; updateCmpBar(); }

function updateCmpBar() {
  const bar = document.getElementById('cmpBar');
  const chips = document.getElementById('cmpChips');
  if (!cmpList.length) { bar.classList.add('hidden'); return; }
  bar.classList.remove('hidden');
  chips.innerHTML = cmpList.map(t =>
    `<div class="cmp-chip">${t} <span class="cmp-chip-x" onclick="removeCmp('${t}')">×</span></div>`
  ).join('');
  document.getElementById('cmpGoBtn').textContent = cmpList.length >= 2 ? T('cmpGoBtn') : (lang==='ar'?'أضف شركة أخرى':'Add 1 more…');
}

function removeCmp(t) { cmpList = cmpList.filter(x => x !== t); updateCmpBar(); }

function showCompare() {
  if (cmpList.length < 2) return;
  getMainTarget().innerHTML = `<div class="loader"><div class="spin"></div></div>`;
  setTimeout(() => renderCompare(cmpList), 200);
}

function renderCompare(tickers) {
  const companies = {};
  tickers.forEach(t => {
    const rows = ANNUAL[t];
    if (!rows?.length) return;
    const m = calcMetrics(rows);
    companies[t] = { rows, m, scores: calcScores(m), stk: STOCK[t]||{} };
  });
  const list = Object.keys(companies);
  if (!list.length) return;
  const isAr = lang === 'ar';

  // Compute sector averages for context
  const allSectors = [...new Set(list.map(t => STOCK[t]?.sector).filter(Boolean))];
  let sectorAvg = null;
  if (allSectors.length === 1) {
    const sec = allSectors[0];
    const secTickers = TICKERS.filter(t => STOCK[t]?.sector === sec && ANNUAL[t]?.length);
    const secScores = secTickers.map(t => calcScores(calcMetrics(ANNUAL[t])));
    const avg = (arr, k) => { const v = arr.map(s => s[k]).filter(Boolean); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
    sectorAvg = { sector: sec, count: secTickers.length, growth: avg(secScores, 'growth'), profitability: avg(secScores, 'profitability'), health: avg(secScores, 'health'), cashflow: avg(secScores, 'cashflow'), overall: avg(secScores, 'overall') };
  }

  // Score table with rankings and optional sector avg row
  const scoreCols = ['growth','profitability','health','cashflow','overall'];
  const scRows = scoreCols.map(k => {
    const vals = list.map(t => companies[t].scores[k] || 0);
    const ranked = [...vals].sort((a, b) => b - a);
    return `<tr>
      <td style="font-weight:600;color:var(--text2);text-transform:capitalize;white-space:nowrap">${T(k)}</td>
      ${list.map((t, i) => {
        const v = companies[t].scores[k];
        const rank = ranked.indexOf(v) + 1;
        const medal = rank === 1 && list.length > 2 ? ' 🥇' : rank === 2 && list.length > 2 ? ' 🥈' : '';
        return `<td style="text-align:center;color:${sCol(v||5)};font-weight:700;font-size:15px">${v||'—'}${medal}</td>`;
      }).join('')}
      ${sectorAvg ? `<td style="text-align:center;color:var(--text3);font-size:13px">${sectorAvg[k]?.toFixed(1) || '—'}</td>` : ''}
    </tr>`;
  }).join('');

  // KPI metrics
  const kpiFields = [{k:'revenue',f:'revenue'},{k:'netIncome',f:'net_income'},{k:'equity',f:'equity'},{k:'ocf',f:'operating_cash_flow'},{k:'fcf',f:'free_cash_flow'},{k:'ebitda',f:'ebitda'}];
  const kpiRows = kpiFields.map(({k,f})=>`<tr>
    <td style="color:var(--text2);white-space:nowrap">${T(k)}</td>
    ${list.map(t=>{const v=companies[t].m.latest[f];return`<td style="text-align:center">${fM(v)}</td>`}).join('')}
    ${sectorAvg ? '<td style="text-align:center;color:var(--text3)">—</td>' : ''}
  </tr>`).join('');

  // Profitability comparison rows
  const profFields = [
    {lbl: isAr?'هامش صافي':'Net Margin', fn: t => { const p = companies[t].m.prof; return p.length ? p[p.length-1].net_margin : null; }},
    {lbl: isAr?'هامش تشغيلي':'Op Margin', fn: t => { const p = companies[t].m.prof; return p.length ? p[p.length-1].op_margin : null; }},
    {lbl: 'ROE', fn: t => { const p = companies[t].m.prof; return p.length ? p[p.length-1].roe : null; }},
    {lbl: 'ROA', fn: t => { const p = companies[t].m.prof; return p.length ? p[p.length-1].roa : null; }},
    {lbl: isAr?'الدين/الملكية':'D/E', fn: t => { const l = companies[t].m.lev; return l.length ? l[l.length-1].dte : null; }},
  ];
  const profRows = profFields.map(pf => {
    const vals = list.map(t => pf.fn(t));
    const valid = vals.filter(v => v != null);
    const best = valid.length ? Math.max(...valid) : null;
    return `<tr><td style="color:var(--text2);white-space:nowrap">${pf.lbl}</td>
      ${list.map((t, i) => {
        const v = vals[i];
        if (v == null) return `<td style="text-align:center">—</td>`;
        const isBest = v === best && valid.length > 1;
        const cls = pf.lbl.includes('D/E') ? '' : pc(v);
        return `<td style="text-align:center;${isBest?'font-weight:700;':''}${cls==='pos'?'color:var(--green)':cls==='neg'?'color:var(--red)':''}">${v.toFixed(1)}${pf.lbl.includes('D/E')?'x':'%'}${isBest?' ★':''}</td>`;
      }).join('')}
      ${sectorAvg ? '<td style="text-align:center;color:var(--text3)">—</td>' : ''}
    </tr>`;
  }).join('');

  const sectorColHdr = sectorAvg ? `<th style="text-align:center;color:var(--text3);font-size:10px">${sectorAvg.sector}<br>${isAr?'متوسط':'Avg'} (${sectorAvg.count})</th>` : '';

  getMainTarget().innerHTML = `
  <div class="cmp-view">
    <div class="cmp-hdr">
      <h2>${T('cmpTitle')}: ${list.join(' vs ')}</h2>
      <p style="color:var(--text2);font-size:13px">${list.length} companies${sectorAvg ? ' · '+sectorAvg.sector+' sector' : ''} · All figures in latest reported year</p>
    </div>
    <div class="tbl-sec">
      <div class="sec-hdr"><div class="sec-title">🎯 ${T('cmpScores')}</div></div>
      <div class="tbl-card"><div class="tbl-scroll"><table>
        <thead><tr><th>Metric</th>${list.map(t=>`<th style="text-align:center">${t}</th>`).join('')}${sectorColHdr}</tr></thead>
        <tbody>${scRows}</tbody>
      </table></div></div>
    </div>
    <div class="tbl-sec">
      <div class="sec-hdr"><div class="sec-title">📊 ${isAr?'مقاييس الربحية':'Profitability Metrics'}</div></div>
      <div class="tbl-card"><div class="tbl-scroll"><table>
        <thead><tr><th>Metric</th>${list.map(t=>`<th style="text-align:center">${t}</th>`).join('')}${sectorColHdr}</tr></thead>
        <tbody>${profRows}</tbody>
      </table></div></div>
    </div>
    <div class="tbl-sec">
      <div class="sec-hdr"><div class="sec-title">💰 ${isAr?'المؤشرات الرئيسية':'Key Metrics'}</div></div>
      <div class="tbl-card"><div class="tbl-scroll"><table>
        <thead><tr><th>Metric</th>${list.map(t=>`<th style="text-align:center">${t}</th>`).join('')}${sectorColHdr}</tr></thead>
        <tbody>${kpiRows}</tbody>
      </table></div></div>
    </div>
    <div class="charts-2">
      <div class="chart-card" style="grid-column:span 2"><div class="chart-title">📈 ${T('revChart')}</div><div class="chart-wrap" style="height:230px"><canvas id="cCmpRev"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">🎯 ${T('cmpScores')}</div><div class="chart-wrap"><canvas id="cCmpScore"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">📊 ${T('yoyChart')}</div><div class="chart-wrap"><canvas id="cCmpYoY"></canvas></div></div>
      <div class="chart-card" style="grid-column:span 2"><div class="chart-title">🎯 ${isAr?'الربحية مقابل النمو':'Profitability vs Growth'}</div><div class="chart-wrap" style="height:230px"><canvas id="cCmpScatter"></canvas></div></div>
    </div>
  </div>`;

  setTimeout(()=>{
    destroyCharts();
    Chart.defaults.color = getChartColors().legend;
    Chart.defaults.borderColor = getChartColors().grid;
    const allYears=[...new Set(list.flatMap(t=>companies[t].rows.map(r=>r.year)))].sort();
    mk('cCmpRev',{type:'line',data:{labels:allYears,datasets:list.map((t,i)=>({
      label:t,data:allYears.map(y=>{const r=companies[t].rows.find(d=>d.year===y);return r?.revenue??null}),
      borderColor:PALETTE[i],backgroundColor:PALETTE[i]+'18',tension:.4,fill:false,pointRadius:4
    }))},options:{responsive:true,maintainAspectRatio:false,scales:scOpts(),plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fM(c.raw)}`}}}}});
    mk('cCmpScore',{type:'bar',data:{labels:list,datasets:[{
      label:T('overallScore'),data:list.map(t=>companies[t].scores.overall||0),
      backgroundColor:list.map((_,i)=>PALETTE[i]+'cc'),borderColor:list.map((_,i)=>PALETTE[i]),borderWidth:1,borderRadius:7
    }]},options:{responsive:true,maintainAspectRatio:false,scales:{...scOpts(),y:{...scOpts().y,max:10}},plugins:{legend:{display:false}}}});
    const yoyYears=[...new Set(list.flatMap(t=>companies[t].m.yoy.map(y=>y.year)))].sort();
    mk('cCmpYoY',{type:'bar',data:{labels:yoyYears,datasets:list.map((t,i)=>({
      label:t,data:yoyYears.map(y=>{const r=companies[t].m.yoy.find(d=>d.year===y);return r?.revenue_growth??null}),
      backgroundColor:PALETTE[i]+'99',borderColor:PALETTE[i],borderWidth:1,borderRadius:4
    }))},options:{responsive:true,maintainAspectRatio:false,scales:scOpts()}});

    // Scatter: Profitability vs Growth
    const scatterData = list.map((t, i) => {
      const c = companies[t];
      const lp = c.m.prof[c.m.prof.length - 1];
      const ly = c.m.yoy[c.m.yoy.length - 1];
      return { x: ly?.revenue_growth ?? 0, y: lp?.net_margin ?? 0, label: t, color: PALETTE[i] };
    });
    mk('cCmpScatter',{type:'scatter',data:{datasets:scatterData.map(d=>({
      label:d.label, data:[{x:d.x,y:d.y}], backgroundColor:d.color+'99', borderColor:d.color, borderWidth:2, pointRadius:10, pointHoverRadius:14
    }))},options:{responsive:true,maintainAspectRatio:false,
      scales:{x:{...scOpts().x,title:{display:true,text:isAr?'نمو الإيرادات %':'Revenue Growth %',color:getChartColors().legend}},y:{...scOpts().y,title:{display:true,text:isAr?'هامش صافي %':'Net Margin %',color:getChartColors().legend}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: Growth ${c.raw.x.toFixed(1)}%, Margin ${c.raw.y.toFixed(1)}%`}}}}});
  },60);
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function doExportCSV(ticker) {
  const rows = ANNUAL[ticker]; if(!rows) return;
  const keys = Object.keys(rows[0]);
  const csv = [keys.join(','), ...rows.map(r=>keys.map(k=>r[k]??'').join(','))].join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv],{type:'text/csv'}));
  a.download = `${ticker}_financials.csv`; a.click();
}

// ── LANGUAGE ──────────────────────────────────────────────────────────────────
function toggleLang() {
  lang = lang==='en'?'ar':'en';
  document.documentElement.dir = lang==='ar'?'rtl':'ltr';
  document.documentElement.lang = lang;
  document.getElementById('langLbl').textContent = T('langLbl');
  document.getElementById('hdrSearch').placeholder = T('hdrSearchPlaceholder');
  document.getElementById('sidebarSearch').placeholder = T('sidebarSearchPlaceholder');
  document.getElementById('sidebarLabel').textContent = T('sidebarLabel');
  document.getElementById('sidebarTickersLabel').textContent = T('sidebarTickersLabel');
  updateCmpBar();
  buildSidebar();
  if (activeTicker) loadTicker(activeTicker);
  else {
    // Refresh welcome/upload text
    if (FILE_LOADED) {
      showLoadedState();
    } else {
      document.getElementById('eyebrowTxt').textContent = T('eyebrowTxt');
      document.getElementById('welcomeH1').innerHTML = T('welcomeH1');
      document.getElementById('welcomeP').textContent = T('welcomeP');
      document.getElementById('hdrSearch').placeholder = lang==='ar'?'ارفع الملف أولاً…':'Upload file first…';
    }
  }
}

// ── FILE HANDLING ─────────────────────────────────────────────────────────────
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('mainDropzone').classList.remove('drag');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

function handleFileSelect() {
  const file = document.getElementById('xlsxInput').files[0];
  if (file) processFile(file);
}

async function processFile(file) {
  const statusEl = document.getElementById('uploadStatus');
  const progressWrap = document.getElementById('progressWrap');
  const progressBar = document.getElementById('progressBar');
  const progressLabel = document.getElementById('progressLabel');
  const progressPct = document.getElementById('progressPct');
  const dropzone = document.getElementById('mainDropzone');

  // Show file name
  const fnEl = document.getElementById('mainFN');
  fnEl.textContent = file.name;
  fnEl.classList.remove('hidden');
  dropzone.classList.add('done');

  // Show progress
  progressWrap.classList.remove('hidden');
  statusEl.classList.add('hidden');

  function setProgress(pct, label) {
    progressBar.style.width = pct + '%';
    progressPct.textContent = pct + '%';
    if (label) progressLabel.textContent = label;
  }

  try {
    setProgress(10, lang === 'ar' ? 'قراءة الملف...' : 'Reading file...');

    const buffer = await file.arrayBuffer();
    setProgress(30, lang === 'ar' ? 'تحليل الأوراق...' : 'Parsing sheets...');

    // Small delay so UI updates
    await new Promise(r => setTimeout(r, 50));

    const result = await parseFinancialXLSX(buffer);
    setProgress(70, lang === 'ar' ? 'بناء قاعدة البيانات...' : 'Building database...');

    await new Promise(r => setTimeout(r, 50));

    DB_RAW = result;
    refreshGlobals();
    FILE_LOADED = true;

    setProgress(90, lang === 'ar' ? 'إعداد الواجهة...' : 'Setting up UI...');
    await new Promise(r => setTimeout(r, 50));

    buildSidebar();
    showLoadedState();

    setProgress(100, lang === 'ar' ? 'تم!' : 'Done!');

    const nTickers = TICKERS.length;
    const nSectors = new Set(TICKERS.map(t => STOCK[t]?.sector).filter(Boolean)).size;
    statusEl.className = 'status ok';
    statusEl.innerHTML = `✓ ${nTickers} ${lang==='ar'?'شركة':'companies'} · ${nSectors} ${lang==='ar'?'قطاعات':'sectors'} loaded from <strong>${file.name}</strong>`;
    statusEl.classList.remove('hidden');

    // Hide progress after a moment
    setTimeout(() => { progressWrap.classList.add('hidden'); }, 1500);

  } catch(err) {
    setProgress(0, '');
    progressWrap.classList.add('hidden');
    statusEl.className = 'status err';
    statusEl.textContent = '✗ Error: ' + (err.message || 'Could not parse file');
    statusEl.classList.remove('hidden');
    console.error(err);
  }
}

function showLoadedState() {
  initChatView();
  // Phase 4: render the alerts panel + subscribe to realtime inserts
  if (typeof updateJournalBadge === 'function') updateJournalBadge().catch(() => {});
  // Portfolio (Commit A)
  if (typeof renderPortfolioPanel === 'function') renderPortfolioPanel();
}

// ── FMP API + IndexedDB CACHE ─────────────────────────────────────────────────
const FMP_BASE = 'https://financialmodelingprep.com/stable';
const IDB_NAME = 'SmartFinAnalyzer';
const IDB_VERSION = 1;
const IDB_STORE = 'companies';
let FMP_API_KEY = '';

// IndexedDB helpers
function openIDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE, { keyPath: 'ticker' });
      }
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = e => reject(e.target.error);
  });
}

async function idbPut(ticker, data) {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put({ ticker, ...data, cachedAt: Date.now() });
    tx.oncomplete = () => resolve();
    tx.onerror = e => reject(e.target.error);
  });
}

async function idbGetAll() {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const req = tx.objectStore(IDB_STORE).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = e => reject(e.target.error);
  });
}

async function idbClear() {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = e => reject(e.target.error);
  });
}

// FMP API fetch helper with rate limiting
async function fmpFetch(endpoint, params, retries = 3) {
  const qs = new URLSearchParams({ ...params, apikey: FMP_API_KEY }).toString();
  const url = `${FMP_BASE}/${endpoint}?${qs}`;
  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = await fetch(url);
    if (res.status === 429) {
      // Rate limited — wait and retry
      const wait = Math.min(5000 * (attempt + 1), 15000);
      console.warn(`FMP 429 rate limit on ${endpoint}/${params.symbol}, waiting ${wait}ms (attempt ${attempt + 1})`);
      await new Promise(r => setTimeout(r, wait));
      continue;
    }
    if (!res.ok) {
      if (attempt < retries) { await new Promise(r => setTimeout(r, 2000)); continue; }
      throw new Error(`FMP ${res.status}: ${res.statusText}`);
    }
    return res.json();
  }
  throw new Error(`FMP: max retries exceeded for ${endpoint}`);
}

// Field mapping: FMP → internal data model
function mapFmpIncomeToInternal(items) {
  return items.filter(r => r.period === 'FY').map(r => ({
    year: parseInt(r.fiscalYear) || new Date(r.date).getFullYear(),
    revenue: r.revenue,
    net_income: r.netIncome,
    operating_income: r.operatingIncome,
    gross_profit: r.grossProfit,
    ebitda: r.ebitda,
    ebit: r.ebit,
    eps_diluted: r.epsDiluted,
    interest_expense: r.interestExpense,
    tax_provision: r.incomeTaxExpense,
    shares_diluted: r.weightedAverageShsOutDil,
    cost_of_revenue: r.costOfRevenue,
    research_and_development: r.researchAndDevelopmentExpenses,
    sga: r.sellingGeneralAndAdministrativeExpenses,
  })).sort((a, b) => a.year - b.year);
}

function mapFmpBalanceToInternal(items) {
  return items.filter(r => r.period === 'FY').map(r => ({
    year: parseInt(r.fiscalYear) || new Date(r.date).getFullYear(),
    total_assets: r.totalAssets,
    total_liabilities: r.totalLiabilities,
    equity: r.totalStockholdersEquity || r.totalEquity,
    total_debt: r.totalDebt,
    cash: r.cashAndCashEquivalents,
    inventory: r.inventory,
    receivables: r.netReceivables || r.accountsReceivables,
    current_assets: r.totalCurrentAssets,
    current_liabilities: r.totalCurrentLiabilities,
    long_term_debt: r.longTermDebt,
    net_ppe: r.propertyPlantEquipmentNet,
    retained_earnings: r.retainedEarnings,
    common_stock: r.commonStock,
  })).sort((a, b) => a.year - b.year);
}

function mapFmpCashFlowToInternal(items) {
  return items.filter(r => r.period === 'FY').map(r => ({
    year: parseInt(r.fiscalYear) || new Date(r.date).getFullYear(),
    operating_cash_flow: r.operatingCashFlow || r.netCashProvidedByOperatingActivities,
    capital_expenditures: r.capitalExpenditure ? -Math.abs(r.capitalExpenditure) : null,
    free_cash_flow: r.freeCashFlow,
    investing_cash_flow: r.netCashProvidedByInvestingActivities,
    financing_cash_flow: r.netCashProvidedByFinancingActivities,
    depreciation: r.depreciationAndAmortization,
    stock_based_comp: r.stockBasedCompensation,
    dividends_paid: r.commonDividendsPaid,
    share_repurchase: r.commonStockRepurchased,
  })).sort((a, b) => a.year - b.year);
}

function mapFmpQuarterlyToInternal(incomeItems, cashItems) {
  const qIncome = (incomeItems || []).filter(r => r.period !== 'FY');
  const qCash = (cashItems || []).filter(r => r.period !== 'FY');
  const dates = new Set([...qIncome.map(r => r.date), ...qCash.map(r => r.date)]);
  const rows = [];
  dates.forEach(dt => {
    const inc = qIncome.find(r => r.date === dt) || {};
    const cf = qCash.find(r => r.date === dt) || {};
    rows.push({
      date: dt,
      revenue: inc.revenue ?? null,
      net_income: inc.netIncome ?? null,
      operating_income: inc.operatingIncome ?? null,
      gross_profit: inc.grossProfit ?? null,
      eps_diluted: inc.epsDiluted ?? null,
      operating_cash_flow: cf.operatingCashFlow ?? cf.netCashProvidedByOperatingActivities ?? null,
      free_cash_flow: cf.freeCashFlow ?? null,
    });
  });
  return rows.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 8);
}

function mapFmpProfileToInternal(profile) {
  if (!profile) return {};
  return {
    price: profile.price,
    marketCap: profile.marketCap,
    pe: profile.price && profile.lastDividend != null ? null : null, // FMP profile doesn't have P/E directly, compute below
    sector: profile.sector || null,
    industry: profile.industry || null,
    currentRatio: null,
    totalDebt: null,
    roe: null,
    roa: null,
    bookValue: null,
    recommendation: null,
    fiftyTwoWeekLow: profile.range ? parseFloat(profile.range.split('-')[0]) : null,
    fiftyTwoWeekHigh: profile.range ? parseFloat(profile.range.split('-')[1]) : null,
    dividendYield: profile.lastDividend && profile.price ? (profile.lastDividend / profile.price) : null,
    beta: profile.beta,
    companyName: profile.companyName,
    exchange: profile.exchange,
    description: profile.description,
    ceo: profile.ceo,
    country: profile.country,
    employees: profile.fullTimeEmployees,
    image: profile.image,
  };
}

// Merge FMP data into internal format matching parseFile output
function buildInternalFromFmp(companyData) {
  const annual = {};
  const quarterly = {};
  const stock = {};

  companyData.forEach(company => {
    const t = company.ticker;
    const inc = company.income || [];
    const bs = company.balance || [];
    const cf = company.cashflow || [];
    const prof = company.profile;

    // Build annual rows (merging income + balance + cashflow by year)
    const incMap = mapFmpIncomeToInternal(inc);
    const bsMap = mapFmpBalanceToInternal(bs);
    const cfMap = mapFmpCashFlowToInternal(cf);

    const allYears = new Set([
      ...incMap.map(r => r.year),
      ...bsMap.map(r => r.year),
      ...cfMap.map(r => r.year)
    ]);

    const rows = [];
    allYears.forEach(yr => {
      const i = incMap.find(r => r.year === yr) || {};
      const b = bsMap.find(r => r.year === yr) || {};
      const c = cfMap.find(r => r.year === yr) || {};

      const entry = {
        year: yr, ticker: t,
        revenue: i.revenue ?? null,
        net_income: i.net_income ?? null,
        operating_income: i.operating_income ?? null,
        gross_profit: i.gross_profit ?? null,
        ebitda: i.ebitda ?? null,
        ebit: i.ebit ?? null,
        eps_diluted: i.eps_diluted ?? null,
        total_assets: b.total_assets ?? null,
        total_liabilities: b.total_liabilities ?? null,
        equity: b.equity ?? null,
        total_debt: b.total_debt ?? null,
        operating_cash_flow: c.operating_cash_flow ?? null,
        capital_expenditures: c.capital_expenditures ?? null,
        free_cash_flow: c.free_cash_flow ?? null,
        interest_expense: i.interest_expense ?? null,
        tax_provision: i.tax_provision ?? null,
        shares_diluted: i.shares_diluted ?? null,
      };
      if (entry.revenue != null || entry.net_income != null) rows.push(entry);
    });

    if (rows.length) annual[t] = rows.sort((a, b) => a.year - b.year);

    // Build quarterly
    const qRows = mapFmpQuarterlyToInternal(inc, cf);
    if (qRows.length) quarterly[t] = qRows;

    // Build stock profile
    const s = mapFmpProfileToInternal(prof);
    // Compute P/E from latest EPS
    const latestInc = incMap[incMap.length - 1];
    if (s.price && latestInc?.eps_diluted) {
      s.pe = s.price / latestInc.eps_diluted;
    }
    // Compute ratios from latest balance sheet
    const latestBs = bsMap[bsMap.length - 1];
    if (latestBs) {
      s.totalDebt = latestBs.total_debt;
      s.bookValue = latestBs.equity;
      if (latestBs.current_assets && latestBs.current_liabilities) {
        s.currentRatio = latestBs.current_assets / latestBs.current_liabilities;
      }
      if (latestInc?.net_income && latestBs.equity && latestBs.equity > 0) {
        s.roe = latestInc.net_income / latestBs.equity;
      }
      if (latestInc?.net_income && latestBs.total_assets && latestBs.total_assets > 0) {
        s.roa = latestInc.net_income / latestBs.total_assets;
      }
    }
    stock[t] = s;
  });

  return { annual, quarterly, stock };
}

// UI helpers
function switchDataTab(tab) {
  document.getElementById('tabXlsx').classList.toggle('active', tab === 'xlsx');
  document.getElementById('tabFmp').classList.toggle('active', tab === 'fmp');
  document.getElementById('dsXlsx').classList.toggle('hidden', tab !== 'xlsx');
  document.getElementById('dsFmp').classList.toggle('hidden', tab !== 'fmp');
}

function setFmpProgress(pct, label) {
  const wrap = document.getElementById('fmpProgressWrap');
  const bar = document.getElementById('fmpProgressBar');
  const lbl = document.getElementById('fmpProgressLabel');
  const pctEl = document.getElementById('fmpProgressPct');
  wrap.classList.remove('hidden');
  bar.style.width = pct + '%';
  lbl.textContent = label;
  pctEl.textContent = Math.round(pct) + '%';
}

function setFmpStatus(type, msg) {
  const el = document.getElementById('fmpStatus');
  el.className = 'fmp-status ' + type;
  el.textContent = msg;
  el.classList.remove('hidden');
}

// Main sync function
async function startFmpSync() {
  const key = document.getElementById('fmpKeyInput').value.trim();
  if (!key) { setFmpStatus('err', 'Please enter your FMP API key'); return; }
  FMP_API_KEY = key;

  const tickerInput = document.getElementById('fmpTickersInput').value.trim();
  let tickers;
  if (tickerInput) {
    tickers = tickerInput.split(/[,\s\n]+/).map(t => t.trim().toUpperCase()).filter(Boolean);
  } else {
    // Default popular tickers
    tickers = ['AAPL','MSFT','GOOGL','AMZN','NVDA','TSLA','META','BRK-B','JPM','JNJ',
      'V','PG','UNH','HD','MA','DIS','PYPL','NFLX','ADBE','CRM','INTC','AMD','QCOM',
      'PEP','KO','MRK','PFE','ABT','TMO','COST','WMT','NKE','MCD','SBUX','BA','CAT',
      'GS','MS','C','AXP','BLK','SPGI','T','VZ','CMCSA','NEE','SO','DUK','XOM','CVX'];
  }

  const btn = document.getElementById('fmpSyncBtn');
  btn.disabled = true;
  btn.textContent = 'Syncing...';
  setFmpStatus('info', `Syncing ${tickers.length} companies from FMP...`);

  const allCompanies = [];
  const errors = [];
  const BATCH_SIZE = 2; // 2 tickers × 6 calls = 12 calls per batch
  const BATCH_DELAY = 3000; // 3s between batches → ~240 calls/min (under 300 limit)
  const totalEst = Math.ceil(tickers.length / BATCH_SIZE) * (BATCH_DELAY / 1000);

  setFmpStatus('info', `Syncing ${tickers.length} companies (~${Math.ceil(totalEst / 60)} min). Don't close this tab.`);

  for (let i = 0; i < tickers.length; i += BATCH_SIZE) {
    const batch = tickers.slice(i, i + BATCH_SIZE);
    const pct = (i / tickers.length) * 100;
    const done = allCompanies.length;
    const errCount = errors.length;
    setFmpProgress(pct, `${done} synced · ${errCount} errors · fetching ${batch.join(', ')}...`);

    const results = await Promise.allSettled(batch.map(async ticker => {
      try {
        const [income, balance, cashflow, profileArr] = await Promise.all([
          fmpFetch('income-statement', { symbol: ticker, period: 'annual', limit: 5 }),
          fmpFetch('balance-sheet-statement', { symbol: ticker, period: 'annual', limit: 5 }),
          fmpFetch('cash-flow-statement', { symbol: ticker, period: 'annual', limit: 5 }),
          fmpFetch('profile', { symbol: ticker }),
        ]);

        // Also fetch quarterly income + cash flow
        const [qIncome, qCash] = await Promise.all([
          fmpFetch('income-statement', { symbol: ticker, period: 'quarter', limit: 8 }).catch(() => []),
          fmpFetch('cash-flow-statement', { symbol: ticker, period: 'quarter', limit: 8 }).catch(() => []),
        ]);

        const profile = Array.isArray(profileArr) ? profileArr[0] : profileArr;
        const data = {
          ticker,
          income: [...(income || []), ...(qIncome || [])],
          balance: balance || [],
          cashflow: [...(cashflow || []), ...(qCash || [])],
          profile: profile || null,
        };

        // Cache in IndexedDB
        await idbPut(ticker, data);
        return data;
      } catch (err) {
        errors.push({ ticker, error: err.message });
        return null;
      }
    }));

    results.forEach(r => { if (r.status === 'fulfilled' && r.value) allCompanies.push(r.value); });

    // Delay between batches to respect 300 calls/min rate limit
    if (i + BATCH_SIZE < tickers.length) await new Promise(r => setTimeout(r, BATCH_DELAY));
  }

  setFmpProgress(100, 'Processing data...');

  if (allCompanies.length === 0) {
    setFmpStatus('err', `Sync failed for all tickers. Check your API key. ${errors.length ? errors[0].error : ''}`);
    btn.disabled = false;
    btn.textContent = 'Sync All';
    return;
  }

  // Build internal data structures — merge with existing cache
  const existingCache = await idbGetAll().catch(() => []);
  // Combine: new data overwrites existing for same ticker, keeps old tickers not in this sync
  const existingMap = {};
  existingCache.forEach(c => { if (!allCompanies.find(n => n.ticker === c.ticker)) existingMap[c.ticker] = c; });
  const mergedCompanies = [...Object.values(existingMap), ...allCompanies];

  const parsed = buildInternalFromFmp(mergedCompanies);
  DB_RAW = parsed;
  refreshGlobals();
  FILE_LOADED = true;

  // Save sync metadata
  try { localStorage.setItem('fmp_last_sync', JSON.stringify({ date: Date.now(), tickers: TICKERS.length, source: 'fmp' })); } catch(e) {}

  setFmpStatus('ok', `✓ Synced ${allCompanies.length} new companies (${TICKERS.length} total in cache, ${errors.length} errors).`);
  setFmpProgress(100, 'Done!');

  btn.disabled = false;
  btn.textContent = 'Re-sync';

  // Activate the app
  buildSidebar();
  showLoadedState();
}

// Auto-load from IndexedDB on page open
async function tryLoadFromCache() {
  try {
    const cached = await idbGetAll();
    if (!cached || cached.length === 0) return false;

    // Check freshness
    const oldest = Math.min(...cached.map(c => c.cachedAt || 0));
    const ageMs = Date.now() - oldest;
    const ageDays = Math.round(ageMs / (1000 * 60 * 60 * 24));

    const parsed = buildInternalFromFmp(cached);
    DB_RAW = parsed;
    refreshGlobals();
    FILE_LOADED = true;

    buildSidebar();
    showLoadedState();

    // Show cache status bar
    const mainArea = document.getElementById('mainArea');
    const cacheBarHtml = `<div class="cache-bar" id="cacheBar">
      <div class="cache-dot ${ageDays < 90 ? 'fresh' : 'stale'}"></div>
      <span>${TICKERS.length} companies from FMP cache (${ageDays} days old)</span>
      <span style="flex:1"></span>
      <button onclick="showDataSourcePage('fmp')">+ Add / Re-sync</button>
      <button onclick="clearCache()">✕ Clear</button>
    </div>`;
    const loadedWelcome = document.getElementById('loadedWelcome');
    if (loadedWelcome) loadedWelcome.insertAdjacentHTML('afterbegin', cacheBarHtml);

    return true;
  } catch (e) {
    console.warn('Cache load failed:', e);
    return false;
  }
}

async function clearCache() {
  await idbClear();
  try { localStorage.removeItem('fmp_last_sync'); } catch(e) {}
  location.reload();
}

// ── LIVE STOCK PRICE (FMP Quote API) ──────────────────────────────────────────
const FMP_QUOTE_KEY = 'MLmSuWJ90zCcDwIwSoktJHsfJaz5Oey0'; // ← Paste your FMP API key here once. It stays in source code.

async function fetchLiveQuote(ticker) {
  if (!FMP_QUOTE_KEY) return null;
  try {
    const res = await fetch(`https://financialmodelingprep.com/stable/quote?symbol=${ticker}&apikey=${FMP_QUOTE_KEY}`);
    if (!res.ok) return null;
    const data = await res.json();
    const q = Array.isArray(data) ? data[0] : data;
    if (!q || !q.price) return null;
    return {
      price: q.price,
      change: q.change,
      // FMP renamed `changesPercentage` → `changePercentage` (no "s"). Accept both for safety.
      changePct: q.changePercentage ?? q.changesPercentage,
      volume: q.volume,
      avgVolume: q.avgVolume,
      marketCap: q.marketCap,
      dayHigh: q.dayHigh,
      dayLow: q.dayLow,
      yearHigh: q.yearHigh,
      yearLow: q.yearLow,
      pe: q.pe,
      open: q.open,
      prevClose: q.previousClose,
    };
  } catch (e) {
    console.warn('Live quote failed for', ticker, e);
    return null;
  }
}

function updateStockBarWithLive(ticker, quote) {
  if (!quote) return;
  const stk = STOCK[ticker];
  if (stk) {
    stk.price = quote.price;
    stk.marketCap = quote.marketCap;
    if (quote.change != null) stk.change = quote.change;
    if (quote.changePct != null) stk.changePct = quote.changePct;
    if (quote.pe) stk.pe = quote.pe;
    if (quote.yearLow) stk.fiftyTwoWeekLow = quote.yearLow;
    if (quote.yearHigh) stk.fiftyTwoWeekHigh = quote.yearHigh;
    stk._priceUpdatedAt = Date.now();
  }

  // Update DOM elements if on this ticker's dashboard
  const priceEl = document.getElementById('livePrice');
  const changeEl = document.getElementById('liveChange');
  const mcapEl = document.getElementById('liveMcap');
  const peEl = document.getElementById('livePe');
  const liveTag = document.getElementById('liveTag');
  const rangeBar = document.getElementById('liveRangeBar');

  if (priceEl && quote.price != null) priceEl.textContent = '$' + quote.price.toFixed(2);
  if (changeEl && quote.change != null && quote.changePct != null) {
    const up = quote.change >= 0;
    changeEl.textContent = `${up ? '+' : ''}${quote.change.toFixed(2)} (${up ? '+' : ''}${quote.changePct.toFixed(2)}%)`;
    changeEl.className = `sb-change ${up ? 'up' : 'down'}`;
  }
  if (mcapEl && quote.marketCap != null) mcapEl.textContent = fM(quote.marketCap);
  if (peEl && quote.pe) peEl.textContent = quote.pe.toFixed(1) + 'x';
  if (liveTag) liveTag.classList.remove('hidden');

  // Update 52-week range bar
  if (rangeBar && quote.yearLow && quote.yearHigh && quote.price) {
    const pct = ((quote.price - quote.yearLow) / (quote.yearHigh - quote.yearLow) * 100).toFixed(0);
    rangeBar.innerHTML = `
      <span style="font-size:13px;color:var(--text2)">$${quote.yearLow.toFixed(2)}</span>
      <div style="flex:1;min-width:80px;height:6px;background:var(--border);border-radius:3px;position:relative">
        <div style="position:absolute;left:${pct}%;top:-3px;width:12px;height:12px;background:var(--accent);border-radius:50%;transform:translateX(-50%);box-shadow:0 0 6px rgba(59,130,246,.5)"></div>
        <div style="height:100%;width:${pct}%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:3px"></div>
      </div>
      <span style="font-size:13px;color:var(--text2)">$${quote.yearHigh.toFixed(2)}</span>
      <span style="font-size:12px;color:var(--text3)">${pct}%</span>`;
  }
}

// ── MARKET MOVERS (gainers, losers, active) ──────────────────────────────────
async function fetchMarketMovers() {
  if (!FMP_QUOTE_KEY) return;
  const el = document.getElementById('marketMovers');
  if (!el) return;

  el.style.display = 'grid';
  el.innerHTML = `<div class="mover-card"><div class="mover-loading">Loading market data...</div></div><div class="mover-card"><div class="mover-loading">Loading...</div></div><div class="mover-card"><div class="mover-loading">Loading...</div></div>`;

  try {
    const [gainers, losers, active] = await Promise.all([
      fetch(`https://financialmodelingprep.com/stable/biggest-gainers?apikey=${FMP_QUOTE_KEY}`).then(r => r.json()).catch(() => []),
      fetch(`https://financialmodelingprep.com/stable/biggest-losers?apikey=${FMP_QUOTE_KEY}`).then(r => r.json()).catch(() => []),
      fetch(`https://financialmodelingprep.com/stable/most-active?apikey=${FMP_QUOTE_KEY}`).then(r => r.json()).catch(() => []),
    ]);

    const renderList = (items, colorFn) => (items || []).slice(0, 8).map(s => {
      const chg = s.changesPercentage || s.change || 0;
      const col = colorFn(chg);
      return `<div class="mover-row" onclick="loadTicker('${s.symbol}')">
        <span class="mover-ticker">${s.symbol}</span>
        <span class="mover-name">${s.name || ''}</span>
        <span class="mover-change" style="color:${col}">${chg > 0 ? '+' : ''}${typeof chg === 'number' ? chg.toFixed(2) : chg}%</span>
      </div>`;
    }).join('') || '<div style="padding:10px;color:var(--text3);font-size:11px">No data</div>';

    el.innerHTML = `
      <div class="mover-card">
        <div class="mover-title gain">🚀 Top Gainers</div>
        ${renderList(gainers, () => 'var(--green)')}
      </div>
      <div class="mover-card">
        <div class="mover-title loss">📉 Top Losers</div>
        ${renderList(losers, () => 'var(--red)')}
      </div>
      <div class="mover-card">
        <div class="mover-title active">⚡ Most Active</div>
        ${renderList(active, c => c >= 0 ? 'var(--green)' : 'var(--red)')}
      </div>`;
  } catch (e) {
    console.warn('Market movers failed:', e);
    el.innerHTML = '<div style="grid-column:span 3;text-align:center;color:var(--text3);font-size:12px;padding:20px">Market data unavailable</div>';
  }
}

// ── STOCK PRICE HISTORY CHART ────────────────────────────────────────────────
async function fetchPriceHistory(ticker, range) {
  if (!FMP_QUOTE_KEY) return [];
  try {
    let url, data;
    const today = new Date();

    if (range === '1d') {
      url = `https://financialmodelingprep.com/stable/historical-chart/5min?symbol=${ticker}&apikey=${FMP_QUOTE_KEY}`;
      const res = await fetch(url);
      if (!res.ok) return [];
      data = await res.json();
      if (!data?.length) return [];
      // Find the most recent trading day in the data (may not be today if market is closed)
      const latestDate = data[0]?.date?.slice(0, 10);
      if (latestDate) data = data.filter(d => d.date?.startsWith(latestDate));
      return data.reverse();

    } else if (range === '5d') {
      // Use 1-hour intervals, filter to last 5 trading days
      url = `https://financialmodelingprep.com/stable/historical-chart/1hour?symbol=${ticker}&apikey=${FMP_QUOTE_KEY}`;
      const res = await fetch(url);
      if (!res.ok) return [];
      data = await res.json();
      // Get unique dates and take only the last 5
      const dates = [...new Set((data || []).map(d => d.date?.slice(0, 10)))].sort().slice(-5);
      data = (data || []).filter(d => dates.includes(d.date?.slice(0, 10)));
      return data.reverse();

    } else {
      // Daily data with date range
      const fromDate = new Date(today);
      if (range === '1m') fromDate.setMonth(fromDate.getMonth() - 1);
      else if (range === '6m') fromDate.setMonth(fromDate.getMonth() - 6);
      else if (range === '1y') fromDate.setFullYear(fromDate.getFullYear() - 1);
      else if (range === '5y') fromDate.setFullYear(fromDate.getFullYear() - 5);
      const from = fromDate.toISOString().slice(0, 10);
      url = `https://financialmodelingprep.com/stable/historical-price-eod/light?symbol=${ticker}&from=${from}&apikey=${FMP_QUOTE_KEY}`;
      const res = await fetch(url);
      if (!res.ok) return [];
      data = await res.json();
      return Array.isArray(data) ? data.reverse() : [];
    }
  } catch (e) {
    console.warn('Price history failed:', e);
    return [];
  }
}

function renderPriceChart(ticker, data, range) {
  const canvas = document.getElementById('priceHistoryChart');
  if (!canvas || !data?.length) {
    const changeEl = document.getElementById('priceChartChange');
    if (changeEl) changeEl.textContent = 'No data';
    return;
  }

  // Destroy existing chart
  if (charts['priceHistory']) { try { charts['priceHistory'].destroy(); } catch(e) {} }

  const labels = data.map(d => {
    if (range === '1d') return d.date?.slice(11, 16) || ''; // HH:MM only
    if (range === '5d') return d.date?.slice(5, 16)?.replace(' ', '\n') || ''; // MM-DD HH:MM
    if (range === '1m' || range === '6m') return d.date?.slice(5, 10) || ''; // MM-DD
    return d.date?.slice(0, 10) || ''; // YYYY-MM-DD for 1Y/5Y
  });
  const prices = data.map(d => d.close || d.price || null);
  const firstPrice = prices.find(p => p != null) || 0;
  const lastPrice = prices[prices.length - 1] || 0;
  const isUp = lastPrice >= firstPrice;
  const color = isUp ? 'rgba(16,185,129,1)' : 'rgba(239,68,68,1)';
  const bgColor = isUp ? 'rgba(16,185,129,0.08)' : 'rgba(239,68,68,0.08)';

  // Update change display
  const changeEl = document.getElementById('priceChartChange');
  if (changeEl) {
    const pctChange = firstPrice ? ((lastPrice - firstPrice) / firstPrice * 100) : 0;
    const dollarChange = lastPrice - firstPrice;
    changeEl.textContent = `${pctChange >= 0 ? '+' : ''}${pctChange.toFixed(2)}% (${dollarChange >= 0 ? '+' : ''}$${Math.abs(dollarChange).toFixed(2)}) · ${range}`;
    changeEl.style.color = isUp ? 'var(--green)' : 'var(--red)';
  }

  // Crosshair plugin — draws vertical + horizontal lines following cursor
  const crosshairPlugin = {
    id: 'crosshair',
    afterDraw(chart) {
      if (!chart._crosshairX) return;
      const { ctx, chartArea: { left, right, top, bottom } } = chart;
      const x = chart._crosshairX;
      const y = chart._crosshairY;

      ctx.save();
      ctx.setLineDash([4, 4]);
      ctx.lineWidth = 1;
      ctx.strokeStyle = theme === 'dark' ? 'rgba(255,255,255,0.3)' : 'rgba(0,0,0,0.2)';

      // Vertical line
      ctx.beginPath();
      ctx.moveTo(x, top);
      ctx.lineTo(x, bottom);
      ctx.stroke();

      // Horizontal line
      if (y >= top && y <= bottom) {
        ctx.beginPath();
        ctx.moveTo(left, y);
        ctx.lineTo(right, y);
        ctx.stroke();
      }

      ctx.restore();
    }
  };

  charts['priceHistory'] = mk('priceHistoryChart', {
    type: 'line',
    data: {
      labels,
      datasets: [{
        data: prices, borderColor: color, backgroundColor: bgColor,
        borderWidth: 2, fill: true, tension: 0.3, pointRadius: 0, pointHoverRadius: 6,
        pointHoverBackgroundColor: color, pointHoverBorderColor: '#fff', pointHoverBorderWidth: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      onHover: (e, elements, chart) => {
        const rect = chart.canvas.getBoundingClientRect();
        chart._crosshairX = e.native ? e.native.offsetX : (e.x - rect.left);
        chart._crosshairY = e.native ? e.native.offsetY : (e.y - rect.top);
        chart.draw();
      },
      scales: {
        x: { display: true, ticks: { maxTicksLimit: 8, font: { size: 9 }, color: theme==='dark'?'rgba(255,255,255,0.4)':'rgba(0,0,0,0.4)' }, grid: { display: false } },
        y: { display: true, position: 'right', ticks: { font: { size: 10 }, color: theme==='dark'?'rgba(255,255,255,0.4)':'rgba(0,0,0,0.4)', callback: v => '$' + v.toFixed(0) }, grid: { color: theme==='dark'?'rgba(255,255,255,0.05)':'rgba(0,0,0,0.05)' } }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: theme==='dark'?'rgba(17,24,39,0.95)':'rgba(255,255,255,0.95)',
          titleColor: theme==='dark'?'#e2e8f0':'#1e293b',
          bodyColor: theme==='dark'?'#94a3b8':'#475569',
          borderColor: theme==='dark'?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',
          borderWidth: 1,
          padding: 10,
          cornerRadius: 8,
          displayColors: false,
          titleFont: { size: 11, weight: '600' },
          bodyFont: { size: 12, family: "'DM Sans', monospace", weight: '700' },
          callbacks: {
            title: (items) => {
              if (!items.length) return '';
              return labels[items[0].dataIndex] || '';
            },
            label: (ctx) => {
              const p = ctx.raw;
              if (p == null) return '';
              const d = data[ctx.dataIndex];
              let lines = [`Price: $${p.toFixed(2)}`];
              if (d?.open) lines.push(`O: $${d.open.toFixed(2)}  H: $${d.high.toFixed(2)}  L: $${d.low.toFixed(2)}`);
              if (d?.volume) lines.push(`Vol: ${d.volume >= 1e6 ? (d.volume/1e6).toFixed(1)+'M' : d.volume >= 1e3 ? (d.volume/1e3).toFixed(0)+'K' : d.volume}`);
              return lines;
            }
          }
        }
      }
    },
    plugins: [crosshairPlugin]
  });
}

async function loadPriceChart(ticker, range) {
  // Update active tab
  document.querySelectorAll('.price-tab').forEach(t => t.classList.remove('active'));
  const activeTab = document.querySelector(`.price-tab[data-range="${range}"]`);
  if (activeTab) activeTab.classList.add('active');

  const canvas = document.getElementById('priceHistoryChart');
  if (canvas) { const ctx = canvas.getContext('2d'); ctx.clearRect(0, 0, canvas.width, canvas.height); }
  const changeEl = document.getElementById('priceChartChange');
  if (changeEl) changeEl.textContent = 'Loading...';

  const data = await fetchPriceHistory(ticker, range);
  renderPriceChart(ticker, data, range);
}

// ── LIVE NEWS (FMP News API) ──────────────────────────────────────────────────
async function fetchLiveNews(ticker) {
  if (!FMP_QUOTE_KEY) return [];
  try {
    const res = await fetch(`https://financialmodelingprep.com/stable/news/stock?symbols=${ticker}&limit=5&apikey=${FMP_QUOTE_KEY}`);
    if (!res.ok) return [];
    return await res.json();
  } catch (e) {
    console.warn('News fetch failed for', ticker, e);
    return [];
  }
}

// Pattern-based pre-classifier — catches obvious cases without an AI call
function quickClassifyHeadline(title, ticker) {
  const t = (title || '').toLowerCase();
  if (!t) return null;

  // Bullish patterns — \W+ allows words between (e.g. "Wins $5B Government Contract")
  const bullPatterns = [
    /\b(?:pops?|jumps?|surges?|soars?|rallies|rockets?|skyrockets?|spikes?)\s+\d+/,
    /\bgains?\s+\d+%/,
    /\bup\s+\d+(?:\.\d+)?%/,
    /\b(?:beats?|tops?|smashes?|crushes?|exceeds?)\s+(?:estimates?|expectations?|forecasts?|consensus)/,
    /\b(?:record|all-time)\s+(?:high|profit|revenue|earnings|sales)/,
    /\bupgraded?\b/,
    /\b(?:approves?|approval|cleared?|authorizes?)\b.*\b(?:fda|drug|treatment|merger)\b/,
    /\bfda\b.*\b(?:approv|cleared|authoriz)/,
    /\b(?:wins?|secures?|lands?|signs?|awarded?)\b[^.]{0,40}\b(?:contract|deal|order|partnership|agreement|funding|grant)\b/,
    /\b(?:raises?|hikes?|boosts?|increases?|lifts?|expands?)\b[^.]{0,30}\b(?:guidance|outlook|forecast|dividend|target|price\s+target|revenue|sales|earnings)\b/,
    /\b(?:announces?|launches?)\s+(?:buyback|share\s+repurchase)/,
    /\bbuyback\b/,
    /\bacquir(?:es?|ed|ing)\s+\w+/,
    /\b(?:bullish|outperform|overweight|strong\s+buy)\s+(?:rating|outlook|call|signal)?/,
    /\bnew\s+(?:high|peak|all-time)/,
    /\bprice\s+target\s+(?:raised|hiked|increased)/,
  ];

  // Bearish patterns
  const bearPatterns = [
    /\b(?:drops?|plunges?|tumbles?|crashes?|sinks?|falls?|slumps?|tanks?|dives?|slides?)\s+\d+/,
    /\bloss(?:es)?\s+\d+%/,
    /\bdown\s+\d+(?:\.\d+)?%/,
    /\b(?:misses?|fails?|disappoints?)\s+(?:estimates?|expectations?|forecasts?|consensus)/,
    /\b(?:cuts?|slashes?|reduces?|lowers?|trims?|slashing)\b[^.]{0,30}\b(?:guidance|outlook|forecast|dividend|target|jobs|workforce|staff|workers|revenue|sales|earnings)\b/,
    /\bdowngraded?\b/,
    /\b(?:lawsuit|sued|investigation|probe|fraud|recall|subpoena|antitrust)\b/,
    /\b(?:layoffs?|firing|fires\s+\d+)\b/,
    /\b(?:bankruptcy|insolvency|chapter\s+11|defaults?)\b/,
    /\b(?:bearish|underperform|underweight|sell)\s+(?:rating|outlook|call|signal)?/,
    /\b(?:crashes?|plunges?|collapses?|tumbles?)\b/,
    /\bnew\s+(?:low|52-week\s+low)\b/,
    /\b(?:warns?|warning)\s+(?:of|on|about)/,
    /\bprice\s+target\s+(?:cut|lowered|reduced)/,
    /\b(?:resign(?:s|ed|ing)?|steps?\s+down|exits?)\b.*\b(?:ceo|cfo|coo|chairman|president)/,
  ];

  const hasBull = bullPatterns.some(re => re.test(t));
  const hasBear = bearPatterns.some(re => re.test(t));
  // Mixed signals (e.g. "AAPL Soars 5%, MSFT Sinks 3%") — defer to AI for ticker-aware judgment
  if (hasBull && hasBear) return null;
  if (hasBull) return 'bullish';
  if (hasBear) return 'bearish';
  return null;
}

async function analyseNewsSentiment(news, ticker) {
  if (!news?.length) return [];
  if (!isAiReady()) return news.map(n => ({ ...n, sentiment: 'unknown', sentimentScore: 0 }));

  // Pre-classify with regex — catches "Pops 22%", "Cuts guidance", etc. without AI
  const preClassified = news.map(n => quickClassifyHeadline(n.title, ticker));
  const ambiguousIdx = preClassified.map((c, i) => c === null ? i : -1).filter(i => i >= 0);
  console.log(`Sentiment: pre-classified ${news.length - ambiguousIdx.length}/${news.length} via patterns, sending ${ambiguousIdx.length} to AI for ${ticker || 'unknown'}`);

  // If everything got pre-classified, skip the AI call entirely
  if (ambiguousIdx.length === 0) {
    return news.map((n, i) => {
      const s = preClassified[i];
      return { ...n, sentiment: s, sentimentScore: s === 'bullish' ? 1 : s === 'bearish' ? -1 : 0 };
    });
  }

  // Build compact item list for the AI: title + short excerpt
  const items = ambiguousIdx.map((origIdx, batchIdx) => {
    const n = news[origIdx];
    const excerpt = (n.text || '').replace(/\s+/g, ' ').trim().slice(0, 250);
    return `${batchIdx + 1}. ${n.title}${excerpt ? `\n   → ${excerpt}` : ''}`;
  }).join('\n\n');

  const subject = ticker || 'the primary stock mentioned';
  const sys = `You are a financial news sentiment analyst. Classify each headline by likely impact on ${subject}'s stock price.

DECISION FRAMEWORK:
- bullish: positive for ${subject} — earnings/revenue beat, contract win, upgrade, FDA approval, product launch, M&A target, positive guidance, sector tailwind helping it, analyst price-target raise, buyback announcement
- bearish: negative for ${subject} — earnings/revenue miss, downgrade, lawsuit, recall, layoffs, executive departure, negative guidance, competitive threat, regulatory action, dilutive offering
- neutral: ONLY for purely informational items with no directional implication — earnings call schedule, "Things to Know Before Earnings", interview transcript, dividend ex-date reminder, generic "trending stock" coverage with no new news

CRITICAL RULES:
- Be DECISIVE. "Neutral" should be uncommon (under 30% of items). Most financial headlines have a directional bias.
- If a title mentions ${subject} alongside other tickers, judge by what's said about ${subject} specifically.
- If a title is about a SECTOR move, classify based on whether ${subject} is in that sector and the direction of the move.
- "Trending Stock: Facts to Know" → neutral. "Stock Pops 22%" → bullish. "Cuts Guidance" → bearish.
- Use the excerpt arrow (→) for additional context if present.

SKIP all reasoning, analysis, or commentary. Output ONLY the JSON array directly. No markdown fences. Start your response with [.

Format: [{"i":1,"s":"bullish"},{"i":2,"s":"neutral"}]`;

  try {
    const response = await fetch(getProxyUrl(), {
      method: 'POST',
      headers: getProxyHeaders(),
      body: JSON.stringify({
        model: DEEPSEEK_MODEL,
        messages: [
          { role: 'system', content: sys },
          { role: 'user', content: `Classify these ${news.length} news items. Respond with the JSON array only:\n${items}` }
        ],
        max_tokens: 2000,
        temperature: 0,
        stream: false
      })
    });

    if (!response.ok) {
      console.warn('Sentiment API error:', response.status);
      return news.map(n => ({ ...n, sentiment: 'neutral', sentimentScore: 0 }));
    }

    const data = await response.json();

    // DeepSeek uses OpenAI format: data.choices[0].message.content
    // V4 Flash sometimes puts output in reasoning_content if it's in reasoning mode
    const msg = data.choices?.[0]?.message;
    let text = '';
    if (msg?.content) {
      text = msg.content;
    } else if (msg?.reasoning_content) {
      // Reasoning mode put output in reasoning_content — try to extract JSON from it
      text = msg.reasoning_content;
    } else if (data.content) {
      text = typeof data.content === 'string' ? data.content : data.content.filter(b => b.type === 'text').map(b => b.text).join('');
    }

    // Clean up — remove markdown fences, whitespace, anything before first [
    text = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    const jsonStart = text.indexOf('[');
    const jsonEnd = text.lastIndexOf(']');
    if (jsonStart >= 0 && jsonEnd > jsonStart) {
      text = text.slice(jsonStart, jsonEnd + 1);
    }

    console.log('Sentiment raw:', text);
    if (!text) {
      console.warn('Sentiment empty — full proxy response:', JSON.stringify(data).slice(0, 800));
      console.warn('Choices[0]:', data.choices?.[0]);
      throw new Error('Empty response from AI');
    }

    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed)) throw new Error('Not an array');

    // AI was given a batch with positions 1..ambiguousIdx.length, where each batch position
    // maps back to ambiguousIdx[batchPos-1] in the original news array.
    return news.map((n, origIdx) => {
      let s = preClassified[origIdx];  // pattern-classified items get their result first
      if (s === null) {
        const batchPos = ambiguousIdx.indexOf(origIdx) + 1;
        const match = parsed.find(p => p.i === batchPos);
        const aiSent = (match?.s || '').toLowerCase();
        s = ['bullish', 'bearish', 'neutral'].includes(aiSent) ? aiSent : 'neutral';
      }
      return { ...n, sentiment: s, sentimentScore: s === 'bullish' ? 1 : s === 'bearish' ? -1 : 0 };
    });

  } catch (e) {
    console.warn('Sentiment analysis failed:', e);
    // Even on AI failure, keep the pattern-classified results — only the ambiguous ones fall back to neutral
    return news.map((n, origIdx) => {
      const s = preClassified[origIdx] || 'neutral';
      return { ...n, sentiment: s, sentimentScore: s === 'bullish' ? 1 : s === 'bearish' ? -1 : 0 };
    });
  }
}

function renderNewsSection(news) {
  const el = document.getElementById('newsFeed');
  if (!el) return;
  if (!news || !news.length) { el.innerHTML = `<div class="news-loading">${lang==='ar'?'لا توجد أخبار متاحة':'No recent news available'}</div>`; return; }

  const isAr = lang === 'ar';

  // Check if sentiment has been scored
  const hasSentiment = news[0]?.sentiment && news[0].sentiment !== 'unknown';
  const isUnknown = news[0]?.sentiment === 'unknown';

  // Sign-in hint when sentiment can't be computed
  let signInHintHtml = '';
  if (isUnknown && !isAiReady()) {
    signInHintHtml = `<div class="sentiment-hint" onclick="showPinModal()">
      <span>🔒</span>
      <span>${isAr ? 'سجّل الدخول لتفعيل تحليل المشاعر بالذكاء الاصطناعي' : 'Sign in to enable AI sentiment analysis on these headlines'}</span>
      <span class="hint-arrow">→</span>
    </div>`;
  }

  // Overall sentiment bar
  let sentimentBarHtml = '';
  if (hasSentiment) {
    const bull = news.filter(n => n.sentiment === 'bullish').length;
    const bear = news.filter(n => n.sentiment === 'bearish').length;
    const neut = news.filter(n => n.sentiment === 'neutral').length;
    const total = news.length;
    const overallScore = news.reduce((s, n) => s + (n.sentimentScore || 0), 0);
    const overallLabel = overallScore > 0 ? (isAr ? 'إيجابي' : 'Bullish') : overallScore < 0 ? (isAr ? 'سلبي' : 'Bearish') : (isAr ? 'محايد' : 'Neutral');
    const overallColor = overallScore > 0 ? 'var(--green)' : overallScore < 0 ? 'var(--red)' : 'var(--text3)';

    sentimentBarHtml = `<div class="sentiment-bar">
      <div class="sentiment-label">${isAr ? 'المزاج العام' : 'Sentiment'}</div>
      <div class="sentiment-meter">
        <div class="sentiment-fill-bull" style="width:${(bull/total*100).toFixed(0)}%"></div>
        <div class="sentiment-fill-neut" style="width:${(neut/total*100).toFixed(0)}%"></div>
        <div class="sentiment-fill-bear" style="width:${(bear/total*100).toFixed(0)}%"></div>
      </div>
      <div class="sentiment-score" style="color:${overallColor}">${overallLabel}</div>
      <div class="sentiment-count">${bull}🟢 ${neut}⚪ ${bear}🔴</div>
    </div>`;
  }

  const sentimentBadge = (s) => {
    if (!s || s === 'unknown') return '';
    const label = s === 'bullish' ? (isAr ? 'إيجابي' : '▲ Bullish') : s === 'bearish' ? (isAr ? 'سلبي' : '▼ Bearish') : (isAr ? 'محايد' : '— Neutral');
    return `<span class="news-sentiment ${s}">${label}</span>`;
  };

  // Stash current news so the Summarize/Chat buttons can reference by index
  _newsItems = news;
  _newsTicker = activeTicker || '';

  const canAnalyze = isAiReady();
  el.innerHTML = signInHintHtml + sentimentBarHtml + news.map((n, i) => `
    <div class="news-item">
      ${n.image ? `<img class="news-img" src="${n.image}" alt="" loading="lazy" onerror="this.style.display='none'"/>` : ''}
      <div class="news-body">
        <a class="news-title" href="${n.url}" target="_blank" rel="noopener">${n.title}</a>
        <div class="news-meta">
          <span class="news-pub">${n.publisher || n.site || ''}</span>
          <span>${n.publishedDate ? new Date(n.publishedDate).toLocaleDateString() : ''}</span>
          ${sentimentBadge(n.sentiment)}
        </div>
        ${canAnalyze ? `<div class="news-actions">
          <button class="news-act-btn" onclick="summarizeArticleAt(${i})">✨ ${isAr?'تلخيص':'Summarize'}</button>
          <button class="news-act-btn" onclick="chatAboutArticleAt(${i})">💬 ${isAr?'ناقش':'Discuss'}</button>
        </div>
        <div class="news-summary-slot" id="newsSummary${i}"></div>` : ''}
      </div>
    </div>`).join('');
}

// ── NEWS: FETCH FULL ARTICLE + SUMMARIZE / DISCUSS ──────────────────────────
let _newsItems = [];
let _newsTicker = '';
const _articleTextCache = new Map(); // url → text (session cache)

function getFetchArticleUrl() { return SUPABASE_URL + '/functions/v1/fetch-article'; }

// Fetch full article text via the edge function (Jina + raw fallback).
// Returns { text, source } or null. Falls back to the FMP excerpt if scrape fails.
async function fetchArticleText(url, fallbackExcerpt = '') {
  if (!url) return fallbackExcerpt ? { text: fallbackExcerpt, source: 'excerpt' } : null;
  if (_articleTextCache.has(url)) return _articleTextCache.get(url);
  try {
    const res = await fetch(getFetchArticleUrl(), {
      method: 'POST',
      headers: getProxyHeaders(),
      body: JSON.stringify({ url })
    });
    const data = await res.json();
    let result;
    if (res.ok && data.text && data.text.length > 200) {
      result = { text: data.text, source: data.source };
    } else if (fallbackExcerpt) {
      result = { text: fallbackExcerpt, source: 'excerpt' };
    } else {
      result = null;
    }
    _articleTextCache.set(url, result);
    return result;
  } catch (e) {
    console.warn('Article fetch failed:', e);
    return fallbackExcerpt ? { text: fallbackExcerpt, source: 'excerpt' } : null;
  }
}

async function summarizeArticleAt(idx) {
  const n = _newsItems[idx];
  if (!n) return;
  const isAr = lang === 'ar';
  const slot = document.getElementById('newsSummary' + idx);
  if (!slot) return;
  // Toggle: if already shown, collapse
  if (slot.dataset.open === '1') { slot.innerHTML = ''; slot.dataset.open = ''; return; }
  slot.dataset.open = '1';
  slot.innerHTML = `<div class="news-summary-loading">⏳ ${isAr?'يجلب المقال ويلخّص...':'Fetching article & summarizing...'}</div>`;

  const article = await fetchArticleText(n.url, (n.text || '').slice(0, 450));
  if (!article?.text) {
    slot.innerHTML = `<div class="news-summary-err">${isAr?'تعذّر جلب المقال':'Could not fetch this article'}</div>`;
    return;
  }

  const ticker = _newsTicker;
  const sys = `You are a financial analyst. Summarize this news article for an investor tracking ${ticker || 'the mentioned company'}.
Output format (concise, use markdown):
**TL;DR:** one sentence.
**Key points:** 2-4 bullets of the concrete facts.
**Why it matters for ${ticker || 'the stock'}:** 1-2 sentences on stock impact (bullish/bearish/neutral and why).
Do NOT invent facts not in the article. Keep it under 130 words total.`;
  const userMsg = `TITLE: ${n.title}\n\nARTICLE${article.source === 'excerpt' ? ' (excerpt only)' : ''}:\n${article.text.slice(0, 4000)}`;

  try {
    const res = await fetch(getProxyUrl(), {
      method: 'POST',
      headers: getProxyHeaders(),
      body: JSON.stringify({
        model: DEEPSEEK_MODEL,
        messages: [{ role: 'system', content: sys }, { role: 'user', content: userMsg }],
        stream: false, temperature: 0.3, max_tokens: 400
      })
    });
    const data = await res.json();
    const msg = data.choices?.[0]?.message;
    const summary = msg?.content || msg?.reasoning_content || '';
    if (!summary) { slot.innerHTML = `<div class="news-summary-err">${isAr?'لم يتم إنشاء ملخص':'No summary generated'}</div>`; return; }
    const srcTag = article.source === 'excerpt'
      ? `<span class="news-src-tag" title="Full article blocked — summarized from excerpt">📄 excerpt</span>`
      : `<span class="news-src-tag" title="Summarized from full article">📰 full article</span>`;
    slot.innerHTML = `<div class="news-summary-box">${srcTag}${fmtAI(summary)}</div>`;
  } catch (e) {
    slot.innerHTML = `<div class="news-summary-err">${e.message}</div>`;
  }
}

// Discuss an article in the chat view — loads it as context, switches to chat.
let _activeArticleContext = null;
async function chatAboutArticleAt(idx) {
  const n = _newsItems[idx];
  if (!n) return;
  const isAr = lang === 'ar';
  const ticker = _newsTicker;

  // Switch to chat view
  if (typeof backToChat === 'function') backToChat();
  addChatMessage('user', `${isAr?'ناقش هذا المقال':'Discuss this article'}: "${n.title}"`);
  const thinkingId = 'artchat_' + Date.now();
  addChatMessage('agent', `<div id="${thinkingId}">⏳ ${isAr?'يجلب المقال...':'Fetching the article...'}</div>`);

  const article = await fetchArticleText(n.url, (n.text || '').slice(0, 450));
  if (!article?.text) {
    const el = document.getElementById(thinkingId);
    if (el) el.innerHTML = isAr?'تعذّر جلب المقال. حاول التلخيص بدلاً من ذلك.':'Could not fetch this article. Try Summarize instead.';
    return;
  }

  // Store context so follow-up questions in chat use this article
  _activeArticleContext = { ticker, title: n.title, url: n.url, text: article.text.slice(0, 4500), source: article.source };
  renderArticleChatChip();

  // Give an opening analysis
  const sys = `You are a financial analyst discussing a news article with an investor tracking ${ticker || 'the mentioned company'}. The user just opened this article to discuss. Give a brief opening (60-90 words): what the article says and the single most important takeaway for ${ticker || 'the stock'}. Then invite follow-up questions. Use markdown. Don't invent facts.`;
  const userMsg = `TITLE: ${n.title}\n\nARTICLE${article.source === 'excerpt' ? ' (excerpt only)' : ''}:\n${article.text.slice(0, 4000)}`;
  try {
    const res = await fetch(getProxyUrl(), {
      method: 'POST', headers: getProxyHeaders(),
      body: JSON.stringify({ model: DEEPSEEK_MODEL, messages: [{ role: 'system', content: sys }, { role: 'user', content: userMsg }], stream: false, temperature: 0.4, max_tokens: 300 })
    });
    const data = await res.json();
    const msg = data.choices?.[0]?.message;
    const opening = msg?.content || msg?.reasoning_content || (isAr?'المقال محمّل. اسأل ما تريد.':'Article loaded. Ask me anything about it.');
    const el = document.getElementById(thinkingId);
    if (el) el.innerHTML = fmtAI(opening);
  } catch (e) {
    const el = document.getElementById(thinkingId);
    if (el) el.innerHTML = isAr?'المقال محمّل. اسأل ما تريد.':'Article loaded. Ask me anything about it.';
  }
  scrollChatToBottom?.();
}

// Answer a follow-up question using the active article as context
async function answerAboutArticle(question) {
  const ctx = _activeArticleContext;
  if (!ctx) return false;
  addChatMessage('user', question);
  const id = 'artans_' + Date.now();
  addChatMessage('agent', `<div id="${id}">⏳</div>`);
  const sys = `You are a financial analyst answering questions about a specific news article the user is discussing (about ${ctx.ticker || 'a company'}). Answer ONLY from the article content provided. If the answer isn't in the article, say so and offer what you can infer. Be concise. Use markdown.`;
  const userMsg = `ARTICLE TITLE: ${ctx.title}\n\nARTICLE CONTENT:\n${ctx.text}\n\n---\nUSER QUESTION: ${question}`;
  try {
    const res = await fetch(getProxyUrl(), {
      method: 'POST', headers: getProxyHeaders(),
      body: JSON.stringify({ model: DEEPSEEK_MODEL, messages: [{ role: 'system', content: sys }, { role: 'user', content: userMsg }], stream: false, temperature: 0.3, max_tokens: 500 })
    });
    const data = await res.json();
    const msg = data.choices?.[0]?.message;
    const ans = msg?.content || msg?.reasoning_content || 'No answer.';
    const el = document.getElementById(id);
    if (el) el.innerHTML = fmtAI(ans);
  } catch (e) {
    const el = document.getElementById(id);
    if (el) el.innerHTML = 'Error: ' + e.message;
  }
  scrollChatToBottom?.();
  return true;
}

function renderArticleChatChip() {
  const area = document.getElementById('chatInputArea');
  if (!area || !_activeArticleContext) return;
  let chip = document.getElementById('articleChatChip');
  if (!chip) {
    chip = document.createElement('div');
    chip.id = 'articleChatChip';
    chip.className = 'article-chat-chip';
    area.insertBefore(chip, area.firstChild);
  }
  const isAr = lang === 'ar';
  chip.innerHTML = `<span class="acc-icon">📰</span>
    <span class="acc-text">${isAr?'تناقش':'Discussing'}: ${_activeArticleContext.title.slice(0, 60)}${_activeArticleContext.title.length > 60 ? '…' : ''}</span>
    <button class="acc-close" onclick="clearArticleChat()" title="${isAr?'إنهاء':'Exit article chat'}">✕</button>`;
}

function clearArticleChat() {
  _activeArticleContext = null;
  document.getElementById('articleChatChip')?.remove();
}

// ── SUPABASE CLOUD DATABASE ───────────────────────────────────────────────────
// Configure these with your Supabase project details
const SUPABASE_URL = 'https://xszfkutbbmavzicvhvgu.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhzemZrdXRiYm1hdnppY3Zodmd1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg0MTgxNzAsImV4cCI6MjA5Mzk5NDE3MH0.9TrQJrb1bvpFCBY_bUSu3uhWXRWn22A0I8clvGiYJ3Y';
let supabaseClient = null;

function getSupabase() {
  if (!supabaseClient && SUPABASE_URL && SUPABASE_ANON_KEY) {
    const creator = (window.supabase && window.supabase.createClient) || (typeof supabase !== 'undefined' && supabase.createClient);
    if (!creator) return null;
    supabaseClient = creator(SUPABASE_URL, SUPABASE_ANON_KEY, { auth: { persistSession: false, autoRefreshToken: false } });
  }
  return supabaseClient;
}

// ── ON-DEMAND TICKER SYNC ───────────────────────────────────────────────────
// Add NEW tickers to the universe from anywhere in the app (header search, console, chat).
// Calls the sync-ticker edge function which fetches FMP → upserts Supabase → returns shaped data
// that we merge directly into local state (no full reload needed).
//
// Examples:
//   syncNewTickers('NBIS')                  → single ticker
//   syncNewTickers(['CEG', 'VST', 'TLN'])   → batch (max 10 per call)
async function syncNewTickers(input) {
  if (!isAiReady()) { console.error('Sign in first'); return null; }
  let list = Array.isArray(input) ? input : [input];
  list = list.map(t => String(t).toUpperCase().trim()).filter(t => /^[A-Z][A-Z.-]{0,5}$/.test(t));
  if (!list.length) { console.error('No valid ticker symbols passed'); return null; }
  // Filter out ones already in the universe to avoid wasted FMP calls
  const newOnly = list.filter(t => !TICKERS.includes(t));
  const skipped = list.filter(t => TICKERS.includes(t));
  if (skipped.length) console.log(`[Sync] Already in universe (skipped): ${skipped.join(', ')}`);
  if (!newOnly.length) { console.log('[Sync] Nothing to do — all tickers already in universe'); return { synced: [], errors: [] }; }

  console.log(`%c[Sync] Fetching ${newOnly.length} new ticker(s) from FMP via edge function...`, 'color:#3b82f6;font-weight:bold');
  const t0 = Date.now();
  try {
    const res = await fetch(getSyncTickerUrl(), {
      method: 'POST',
      headers: getProxyHeaders(),
      body: JSON.stringify({ tickers: newOnly })
    });
    const data = await res.json();
    if (!res.ok) {
      console.error('[Sync] FAILED:', data);
      return data;
    }
    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    console.log(`%c[Sync] Done in ${elapsed}s — ${data.count} synced, ${data.errors?.length || 0} errors`, 'color:#10b981;font-weight:bold');
    if (data.errors?.length) console.warn('[Sync] Errors:', data.errors);

    // Merge synced tickers into local state
    for (const item of (data.synced || [])) {
      mergeTickerLocally(item);
      console.log(`[Sync] ✓ ${item.ticker} merged into local universe`);
    }
    // Re-render sidebar so new tickers are immediately visible
    if (data.synced?.length && typeof buildSidebar === 'function') buildSidebar();
    return data;
  } catch (e) {
    console.error('[Sync] Network error:', e);
    return { error: e.message };
  }
}

// Merge a single synced ticker's data into the live STOCK/ANNUAL/QUARTERLY/ESTIMATES/TICKERS arrays
function mergeTickerLocally(item) {
  if (!item?.ticker) return;
  const t = item.ticker;
  // Annual: convert DB column names to in-memory format used by calcMetrics
  ANNUAL[t] = (item.annual || []).map(r => ({
    year: r.fiscal_year, ticker: t,
    revenue: r.revenue, net_income: r.net_income, operating_income: r.operating_income,
    gross_profit: r.gross_profit, ebitda: r.ebitda, ebit: r.ebit,
    eps_diluted: r.eps_diluted, total_assets: r.total_assets,
    total_liabilities: r.total_liabilities, equity: r.equity,
    total_debt: r.total_debt, operating_cash_flow: r.operating_cash_flow,
    capital_expenditures: r.capital_expenditures, free_cash_flow: r.free_cash_flow,
    interest_expense: r.interest_expense, tax_provision: r.tax_provision,
    shares_diluted: r.shares_diluted,
  })).sort((a, b) => a.year - b.year);

  // Quarterly
  QUARTERLY[t] = (item.quarterly || []).map(r => ({
    date: r.fiscal_date, revenue: r.revenue, net_income: r.net_income,
    operating_income: r.operating_income, gross_profit: r.gross_profit,
    eps_diluted: r.eps_diluted, operating_cash_flow: r.operating_cash_flow,
    free_cash_flow: r.free_cash_flow,
  }));

  // Estimates (only if ESTIMATES global exists)
  if (typeof ESTIMATES !== 'undefined' && item.estimates) {
    ESTIMATES[t] = item.estimates.map(e => ({
      date: e.date, revenueAvg: e.revenue_avg, revenueLow: e.revenue_low, revenueHigh: e.revenue_high,
      netIncomeAvg: e.net_income_avg, epsAvg: e.eps_avg, epsLow: e.eps_low, epsHigh: e.eps_high,
      numAnalystsRevenue: e.num_analysts_revenue, numAnalystsEps: e.num_analysts_eps,
    }));
  }

  // Profile (build the STOCK object the same way loadFromSupabase does)
  const p = item.profile;
  if (p) {
    const latestAnnual = ANNUAL[t]?.[ANNUAL[t].length - 1];
    const s = {
      price: p.price, marketCap: p.market_cap, sector: p.sector,
      industry: p.industry, beta: p.beta, companyName: p.company_name,
      exchange: p.exchange, country: p.country, description: p.description,
      ceo: p.ceo, employees: p.employees, image: p.image, pe: null,
      fiftyTwoWeekLow: p.range_52w ? parseFloat(String(p.range_52w).split('-')[0]) : null,
      fiftyTwoWeekHigh: p.range_52w ? parseFloat(String(p.range_52w).split('-')[1]) : null,
      dividendYield: p.last_dividend && p.price ? (p.last_dividend / p.price) : null,
    };
    if (s.price && latestAnnual?.eps_diluted) s.pe = s.price / latestAnnual.eps_diluted;
    STOCK[t] = s;
  }

  if (!TICKERS.includes(t)) TICKERS.push(t);
}

// ── BULK CLOUD REFRESH ────────────────────────────────────────────────────────
// Owner action: re-sync ALL companies' financials (annual + QUARTERLY + profile +
// estimates) from FMP into Supabase via the sync-ticker edge function, 10 at a time.
// Unlike syncNewTickers (which skips tickers already in the universe), this
// intentionally re-syncs existing tickers so their quarterly data is refreshed.
let _bulkRefreshAbort = false;

async function refreshAllCloudData() {
  if (!isAiReady()) { showPinModal(); return; }
  const all = [...TICKERS].sort();
  if (!all.length) { alert(lang === 'ar' ? 'لا توجد شركات محمّلة بعد.' : 'No companies loaded yet.'); return; }

  const BATCH = 10; // sync-ticker caps at 10/call to fit Supabase's 50s edge timeout
  const totalBatches = Math.ceil(all.length / BATCH);
  const estMin = Math.max(1, Math.round(totalBatches * 30 / 60)); // ~30s per batch

  const go = confirm(lang === 'ar'
    ? `تحديث البيانات المالية (بما فيها الفصلية) لكل الشركات (${all.length}) من FMP إلى قاعدة بيانات Supabase.\n\nقد يستغرق ~${estMin} دقيقة ويستهلك حصة FMP الخاصة بك. أبقِ هذا التبويب مفتوحاً.\n\nهل تريد المتابعة؟`
    : `Refresh financials (incl. quarterly) for all ${all.length} companies from FMP into your Supabase database.\n\nThis takes ~${estMin} min and uses your FMP API quota. Keep this tab open.\n\nProceed?`);
  if (!go) return;

  await runBulkCloudRefresh(all);
}

// Refresh only the ticker(s) typed into the sidebar input (comma/space separated).
// Handy for testing a single name without a full ~20-30 min run.
async function refreshSomeCloudData() {
  if (!isAiReady()) { showPinModal(); return; }
  const inp = document.getElementById('refreshTickersInput');
  const list = [...new Set(((inp && inp.value) || '')
    .split(/[,\s]+/).map(t => t.trim().toUpperCase())
    .filter(t => /^[A-Z][A-Z.-]{0,5}$/.test(t)))];
  if (!list.length) {
    alert(lang === 'ar' ? 'اكتب رمزاً واحداً أو أكثر (مثال: AAPL, MSFT).' : 'Enter one or more tickers (e.g. AAPL, MSFT).');
    return;
  }
  await runBulkCloudRefresh(list);
  if (inp) inp.value = '';
}

// Shared engine: sync a list of tickers through the sync-ticker edge function,
// 10 at a time, with a progress modal. Used by both "refresh all" and "refresh some".
async function runBulkCloudRefresh(list) {
  const BATCH = 10; // sync-ticker caps at 10/call to fit Supabase's 50s edge timeout
  _bulkRefreshAbort = false;
  const ui = openBulkRefreshModal(list.length);

  let done = 0, synced = 0, failed = 0;
  const errors = [];

  for (let i = 0; i < list.length; i += BATCH) {
    if (_bulkRefreshAbort) break;
    const batch = list.slice(i, i + BATCH);
    ui.update(done, synced, failed, (lang === 'ar' ? 'جارٍ مزامنة: ' : 'Syncing: ') + batch.join(', '));

    try {
      const res = await fetch(getSyncTickerUrl(), {
        method: 'POST',
        headers: getProxyHeaders(),
        body: JSON.stringify({ tickers: batch }),
      });
      const data = await res.json().catch(() => ({}));

      if (res.ok) {
        (data.synced || []).forEach(item => { try { mergeTickerLocally(item); } catch (e) {} synced++; });
        (data.errors || []).forEach(e => { errors.push(e); failed++; });
      } else {
        batch.forEach(t => { errors.push({ ticker: t, error: data.error || ('HTTP ' + res.status) }); failed++; });
        if (res.status === 401 || res.status === 403) {
          ui.finish(synced, failed, errors, lang === 'ar' ? 'انتهت الجلسة — سجّل الدخول مجدداً.' : 'Session expired — please sign in again.');
          return;
        }
      }
    } catch (e) {
      batch.forEach(t => { errors.push({ ticker: t, error: e.message }); failed++; });
    }

    done += batch.length;
    ui.update(done, synced, failed, '');
  }

  // Reflect refreshed data immediately. mergeTickerLocally already updated the live
  // globals in place; do NOT call refreshGlobals() here — it rebuilds from the stale
  // DB_RAW and would discard everything we just merged.
  if (typeof buildSidebar === 'function') buildSidebar();

  ui.finish(synced, failed, errors,
    _bulkRefreshAbort
      ? (lang === 'ar' ? 'أُلغيت المزامنة.' : 'Cancelled.')
      : (lang === 'ar' ? '✓ اكتمل التحديث.' : '✓ Refresh complete.'));
}

function openBulkRefreshModal(total) {
  document.getElementById('bulkRefreshOverlay')?.remove();
  const isAr = lang === 'ar';
  const overlay = document.createElement('div');
  overlay.id = 'bulkRefreshOverlay';
  overlay.className = 'pin-overlay';
  overlay.innerHTML = `<div class="pin-card" style="max-width:460px;text-align:left">
    <h2 style="margin-bottom:4px">🔄 ${isAr ? 'تحديث بيانات السحابة' : 'Refresh Cloud Data'}</h2>
    <p style="margin-bottom:14px;color:var(--text2)">${isAr ? `مزامنة ${total} شركة من FMP إلى Supabase…` : `Syncing ${total} companies from FMP → Supabase…`}</p>
    <div style="height:10px;background:rgba(148,163,184,.25);border-radius:6px;overflow:hidden;margin-bottom:8px">
      <div id="brBar" style="height:100%;width:0%;background:linear-gradient(90deg,#10b981,#3b82f6);transition:width .3s"></div>
    </div>
    <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--text2)">
      <span id="brCounts">0 / ${total}</span><span id="brPct">0%</span>
    </div>
    <div id="brNote" style="font-size:11px;color:var(--text3);margin-top:8px;min-height:16px;word-break:break-word"></div>
    <div id="brActions" style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end">
      <button id="brCancel" class="pin-submit" style="background:var(--red);width:auto;padding:8px 16px" onclick="_bulkRefreshAbort=true;this.disabled=true;this.textContent='${isAr ? 'يتم الإيقاف…' : 'Stopping…'}'">${isAr ? 'إيقاف' : 'Stop'}</button>
    </div>
  </div>`;
  document.body.appendChild(overlay);

  return {
    update(d, ok, fail, note) {
      const pct = total ? Math.round((d / total) * 100) : 0;
      const bar = document.getElementById('brBar'); if (bar) bar.style.width = pct + '%';
      const c = document.getElementById('brCounts'); if (c) c.textContent = `${d} / ${total}  ·  ${ok} ✓  ·  ${fail} ✕`;
      const p = document.getElementById('brPct'); if (p) p.textContent = pct + '%';
      if (note) { const n = document.getElementById('brNote'); if (n) n.textContent = note; }
    },
    finish(ok, fail, errs, msg) {
      const bar = document.getElementById('brBar'); if (bar) bar.style.width = '100%';
      const isAr2 = lang === 'ar';
      const n = document.getElementById('brNote');
      if (n) {
        let html = `<strong style="color:var(--green)">${msg}</strong><br>${isAr2 ? 'تمّت' : 'Synced'} ${ok} · ${isAr2 ? 'أخطاء' : 'errors'} ${fail}`;
        if (errs && errs.length) {
          const sample = errs.slice(0, 8).map(e => `${e.ticker}: ${e.error}`).join('<br>');
          html += `<details style="margin-top:8px"><summary style="cursor:pointer">${isAr2 ? 'عرض الأخطاء' : 'Show errors'} (${errs.length})</summary><div style="margin-top:6px;font-size:10px;color:var(--text3);max-height:120px;overflow:auto">${sample}${errs.length > 8 ? '<br>…' : ''}</div></details>`;
        }
        n.innerHTML = html;
      }
      const acts = document.getElementById('brActions');
      if (acts) acts.innerHTML = `<button class="pin-submit" style="width:auto;padding:8px 16px" onclick="document.getElementById('bulkRefreshOverlay')?.remove()">${isAr2 ? 'إغلاق' : 'Close'}</button>`;
    }
  };
}

async function loadFromSupabase() {
  const sb = getSupabase();
  if (!sb) return false;

  // Show loading state on the welcome screen
  const uw = document.getElementById('uploadWelcome');
  if (uw) {
    uw.innerHTML = `<div class="welcome-eyebrow">☁️ Loading from cloud database...</div>
      <h1 style="font-size:clamp(28px,3.5vw,44px)">Connecting to Supabase</h1>
      <p>Fetching financial data for all companies...</p>
      <div style="margin-top:20px"><div class="loader"><div class="spin"></div></div></div>`;
  }

  try {
    // Check if database has data
    const { data: meta, error: metaErr } = await sb.from('sync_metadata').select('*').eq('id', 1).single();
    if (metaErr) { console.warn('sync_metadata error:', metaErr.message); }

    // Helper: fetch all rows with pagination (Supabase default limit = 1000)
    async function fetchAll(table, orderCol, ascending) {
      const PAGE = 1000;
      let all = [], from = 0;
      while (true) {
        const { data, error } = await sb.from(table).select('*').order(orderCol, { ascending }).range(from, from + PAGE - 1);
        if (error) throw error;
        if (!data || data.length === 0) break;
        all = all.concat(data);
        if (data.length < PAGE) break;
        from += PAGE;
      }
      return all;
    }

    // Fetch all tables in PARALLEL (much faster than sequential)
    const updateProgress = (msg) => { if (uw) { const p = uw.querySelector('p'); if (p) p.textContent = msg; } };

    updateProgress('Fetching company data...');
    const [profiles, annuals, quarters, estimates] = await Promise.all([
      fetchAll('company_profiles', 'symbol', true),
      fetchAll('annual_financials', 'fiscal_year', true),
      fetchAll('quarterly_financials', 'fiscal_date', false),
      fetchAll('analyst_estimates', 'date', true).catch(() => []),
    ]);

    if (!profiles?.length) { console.warn('No profiles in Supabase'); return false; }
    updateProgress(`Processing ${profiles.length} companies...`);

    console.log(`Supabase: ${profiles.length} profiles, ${annuals.length} annual, ${quarters.length} quarterly, ${estimates.length} estimates`);

    // Build internal data structures
    const annual = {};
    const quarterly = {};
    const stock = {};

    (annuals || []).forEach(r => {
      if (!annual[r.symbol]) annual[r.symbol] = [];
      annual[r.symbol].push({
        year: r.fiscal_year, ticker: r.symbol,
        revenue: r.revenue, net_income: r.net_income, operating_income: r.operating_income,
        gross_profit: r.gross_profit, ebitda: r.ebitda, ebit: r.ebit,
        eps_diluted: r.eps_diluted, total_assets: r.total_assets,
        total_liabilities: r.total_liabilities, equity: r.equity,
        total_debt: r.total_debt, operating_cash_flow: r.operating_cash_flow,
        capital_expenditures: r.capital_expenditures, free_cash_flow: r.free_cash_flow,
        interest_expense: r.interest_expense, tax_provision: r.tax_provision,
        shares_diluted: r.shares_diluted,
      });
    });

    (quarters || []).forEach(r => {
      if (!quarterly[r.symbol]) quarterly[r.symbol] = [];
      quarterly[r.symbol].push({
        date: r.fiscal_date, revenue: r.revenue, net_income: r.net_income,
        operating_income: r.operating_income, gross_profit: r.gross_profit,
        eps_diluted: r.eps_diluted, operating_cash_flow: r.operating_cash_flow,
        free_cash_flow: r.free_cash_flow,
      });
    });

    // Group estimates by symbol
    const estBySymbol = {};
    (estimates || []).forEach(e => {
      if (!estBySymbol[e.symbol]) estBySymbol[e.symbol] = [];
      estBySymbol[e.symbol].push({
        date: e.date, revenueAvg: e.revenue_avg, revenueLow: e.revenue_low, revenueHigh: e.revenue_high,
        netIncomeAvg: e.net_income_avg, epsAvg: e.eps_avg, epsLow: e.eps_low, epsHigh: e.eps_high,
        numAnalystsRevenue: e.num_analysts_revenue, numAnalystsEps: e.num_analysts_eps,
      });
    });

    (profiles || []).forEach(p => {
      const t = p.symbol;
      const latestAnnual = annual[t]?.[annual[t].length - 1];
      const s = {
        price: p.price, marketCap: p.market_cap, sector: p.sector,
        industry: p.industry, beta: p.beta, companyName: p.company_name,
        exchange: p.exchange, country: p.country, description: p.description,
        ceo: p.ceo, employees: p.employees, image: p.image, pe: null,
        fiftyTwoWeekLow: p.range_52w ? parseFloat(p.range_52w.split('-')[0]) : null,
        fiftyTwoWeekHigh: p.range_52w ? parseFloat(p.range_52w.split('-')[1]) : null,
        dividendYield: p.last_dividend && p.price ? (p.last_dividend / p.price) : null,
      };
      if (s.price && latestAnnual?.eps_diluted) s.pe = s.price / latestAnnual.eps_diluted;
      if (latestAnnual?.net_income && latestAnnual?.equity && latestAnnual.equity > 0) s.roe = latestAnnual.net_income / latestAnnual.equity;
      if (latestAnnual?.net_income && latestAnnual?.total_assets && latestAnnual.total_assets > 0) s.roa = latestAnnual.net_income / latestAnnual.total_assets;
      if (latestAnnual?.current_assets && latestAnnual?.total_liabilities) s.currentRatio = latestAnnual.current_assets / (latestAnnual.total_liabilities || 1);
      stock[t] = s;
    });

    // Activate
    DB_RAW = { annual, quarterly, stock, estimates: estBySymbol };
    refreshGlobals();
    FILE_LOADED = true;

    buildSidebar();
    showLoadedState();

    // Show cloud status bar
    const syncDate = meta?.last_sync ? new Date(meta.last_sync).toLocaleDateString() : 'unknown';
    const loadedWelcome = document.getElementById('loadedWelcome');
    if (loadedWelcome) {
      loadedWelcome.insertAdjacentHTML('afterbegin', `<div class="cache-bar" id="cacheBar">
        <div class="cache-dot fresh"></div>
        <span>☁️ ${TICKERS.length} companies loaded from cloud database (synced ${syncDate})</span>
        <span style="flex:1"></span>
      </div>`);
    }

    console.log(`Loaded ${TICKERS.length} companies from Supabase`);
    return true;
  } catch (err) {
    console.error('Supabase load failed:', err);
    // Show error on welcome screen
    if (uw) {
      uw.innerHTML = `<div class="welcome-eyebrow" style="color:var(--red)">❌ Cloud connection failed</div>
        <h1 style="font-size:clamp(22px,3vw,36px)">Could not load data</h1>
        <p style="color:var(--text2)">${err.message || 'Unknown error'}</p>
        <p style="color:var(--text3);font-size:13px;margin-top:12px">Check your Supabase configuration or try refreshing the page.</p>`;
    }
    return false;
  }
}

function showDataSourcePage(tab) {
  const ws = document.getElementById('welcomeScreen');
  const mainArea = document.getElementById('mainArea');

  // Hide dashboards and show welcome
  const dfv = document.getElementById('dashFullView');
  if (dfv) dfv.classList.add('hidden');
  const cv = document.getElementById('chatView');
  if (cv) cv.classList.add('hidden');
  const dp = document.getElementById('dashPanel');
  if (dp) dp.classList.add('hidden');
  mainArea.classList.remove('dash-open');

  // Show the upload welcome (with tabs), hide the loaded welcome
  ws.classList.remove('hidden');
  const uw = document.getElementById('uploadWelcome');
  const lw = document.getElementById('loadedWelcome');
  if (uw) uw.classList.remove('hidden');
  if (lw) lw.classList.add('hidden');

  // Switch to requested tab
  if (tab) switchDataTab(tab);

  // Pre-fill info about existing cache
  if (tab === 'fmp' && TICKERS.length) {
    const input = document.getElementById('fmpTickersInput');
    if (input && !input.value.trim()) {
      input.placeholder = `Currently cached: ${TICKERS.length} tickers (${TICKERS.slice(0, 10).join(', ')}${TICKERS.length > 10 ? '...' : ''})\nEnter tickers to add or re-sync.`;
    }
    // Restore API key if we have it
    const keyInput = document.getElementById('fmpKeyInput');
    if (keyInput && FMP_API_KEY && !keyInput.value) keyInput.value = FMP_API_KEY;
  }

  // Clear active ticker state
  activeTicker = null;
}

// ── AUTH + PERSISTENT MEMORY (Shared Password + Supabase) ─────────────────────
let USER_AUTH_HASH = null;
let memorySaveTimeout = null;

async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

function showPinModal() {
  const existing = document.getElementById('pinOverlay');
  if (existing) existing.remove();
  const isAr = lang === 'ar';
  const isLoggedIn = !!USER_AUTH_HASH;

  const overlay = document.createElement('div');
  overlay.id = 'pinOverlay';
  overlay.className = 'pin-overlay';
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="pin-card">
    <h2>${isLoggedIn ? (isAr ? '✓ تم تسجيل الدخول' : '✓ Signed In') : (isAr ? '🔐 تسجيل الدخول' : '🔐 Sign In')}</h2>
    <p>${isLoggedIn
      ? (isAr ? 'أنت مسجل الدخول. بياناتك محفوظة ومتزامنة.' : 'You are signed in. Your data is saved and synced.')
      : (isAr ? 'أدخل كلمة المرور للوصول إلى التحليل الذكي والذاكرة.' : 'Enter the password to access AI analysis and synced memory.')}</p>
    ${isLoggedIn ? '' : `<input class="pin-input" id="pinInput" type="password" placeholder="${isAr ? 'كلمة المرور' : 'Password'}" onkeydown="if(event.key==='Enter')submitPassword()"/>
    <button class="pin-submit" onclick="submitPassword()">${isAr ? 'دخول' : 'Sign In'}</button>`}
    ${isLoggedIn ? `<button class="pin-submit" style="background:var(--red)" onclick="signOut()">${isAr ? 'تسجيل خروج' : 'Sign Out'}</button>` : ''}
    <div id="pinStatus"></div>
    <div class="pin-note">${isAr ? '🔒 كلمة المرور مطلوبة لاستخدام التحليل الذكي.' : '🔒 Password is required to use AI analysis features.'}</div>
  </div>`;
  document.body.appendChild(overlay);
  const inp = document.getElementById('pinInput');
  if (inp) inp.focus();
}

async function submitPassword() {
  const inp = document.getElementById('pinInput');
  const statusEl = document.getElementById('pinStatus');
  if (!inp || !inp.value.trim()) { if (statusEl) { statusEl.className = 'pin-status err'; statusEl.textContent = lang === 'ar' ? 'أدخل كلمة المرور' : 'Please enter the password'; } return; }

  const password = inp.value.trim();
  if (statusEl) { statusEl.className = 'pin-status info'; statusEl.textContent = lang === 'ar' ? 'جاري التحقق...' : 'Verifying...'; }

  const hash = await hashPassword(password);

  try {
    // Auth-only check: send an empty-messages request so the proxy validates the password
    // BUT short-circuits at the "messages required" check before calling OpenAI/DeepSeek.
    // Saves 1-3 seconds vs. making a real AI ping just to verify the password.
    const res = await fetch(getProxyUrl(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${SUPABASE_ANON_KEY}`, 'x-auth-hash': hash },
      body: JSON.stringify({}),
    });

    if (res.status === 401 || res.status === 403) {
      if (statusEl) { statusEl.className = 'pin-status err'; statusEl.textContent = lang === 'ar' ? 'كلمة المرور غير صحيحة' : 'Incorrect password. Try again.'; }
      return;
    }
    // 400 = auth OK, missing messages (expected); 200 = auth OK with real response. Either means success.

    USER_AUTH_HASH = hash;
    try { localStorage.setItem('sfa_auth_hash', USER_AUTH_HASH); } catch(e) {}

    await loadMemoryFromCloud();
    updateMemoryBadge();

    if (statusEl) { statusEl.className = 'pin-status ok'; statusEl.textContent = lang === 'ar' ? '✓ تم تسجيل الدخول بنجاح' : '✓ Signed in successfully!'; }
    setTimeout(() => { const o = document.getElementById('pinOverlay'); if (o) o.remove(); }, 1200);
  } catch (e) {
    if (statusEl) { statusEl.className = 'pin-status err'; statusEl.textContent = lang === 'ar' ? 'خطأ في الاتصال' : 'Connection error. Try again.'; }
  }
}

function signOut() {
  USER_AUTH_HASH = null;
  agentMemory = {};
  watchlist = [];
  try { localStorage.removeItem('sfa_auth_hash'); } catch(e) {}
  renderWatchlistPanel();
  updateMemoryBadge();
  const overlay = document.getElementById('pinOverlay');
  if (overlay) overlay.remove();
}

function updateMemoryBadge() {
  const btn = document.getElementById('memoryBtn');
  const lbl = document.getElementById('memoryLbl');
  if (!btn || !lbl) return;
  if (USER_AUTH_HASH) {
    lbl.innerHTML = `<span class="memory-badge">SIGNED IN</span>`;
    btn.style.borderColor = 'rgba(16,185,129,.3)';
    btn.style.color = 'var(--green)';
  } else {
    lbl.textContent = lang === 'ar' ? 'دخول' : 'Sign In';
    btn.style.borderColor = '';
    btn.style.color = '';
  }
}

// Cloud memory operations — uses a fixed shared identity for all authenticated users
const SHARED_MEMORY_ID = 'shared_user';

async function loadMemoryFromCloud() {
  if (!USER_AUTH_HASH) return false;
  const sb = getSupabase();
  if (!sb) return false;

  try {
    const { data, error } = await sb.from('user_memory').select('key, value').eq('pin_hash', SHARED_MEMORY_ID);
    if (error) { console.warn('Memory load error:', error.message); return false; }
    if (!data || !data.length) return false;

    data.forEach(row => {
      if (row.key === 'watchlist') watchlist = row.value || [];
      if (row.key === 'portfolio') portfolio = row.value || [];
      if (row.key === 'investor_profile' && row.value) investorProfile = { ...investorProfile, ...row.value };
      if (row.key === 'agent_memory') agentMemory = row.value || {};
      if (row.key === 'settings') {
        const s = row.value || {};
        // Only restore the saved model if it's still in the current AI_MODELS list
        // (prevents stale model IDs from past versions overriding the current default)
        if (s.deepseek_model && AI_MODELS.some(m => m.id === s.deepseek_model)) {
          DEEPSEEK_MODEL = s.deepseek_model;
        }
      }
    });

    renderWatchlistPanel();
    const modelBtn = document.getElementById('modelBtn');
    if (modelBtn) modelBtn.textContent = getModelMeta(DEEPSEEK_MODEL).label;
    return true;
  } catch (e) {
    console.warn('Memory load failed:', e);
    return false;
  }
}

async function saveMemoryToCloud() {
  if (!USER_AUTH_HASH) return;
  const sb = getSupabase();
  if (!sb) return;

  try {
    const now = new Date().toISOString();
    await sb.from('user_memory').upsert([
      { pin_hash: SHARED_MEMORY_ID, key: 'watchlist', value: watchlist, updated_at: now },
      { pin_hash: SHARED_MEMORY_ID, key: 'portfolio', value: portfolio, updated_at: now },
      { pin_hash: SHARED_MEMORY_ID, key: 'investor_profile', value: investorProfile, updated_at: now },
      { pin_hash: SHARED_MEMORY_ID, key: 'agent_memory', value: agentMemory, updated_at: now },
      { pin_hash: SHARED_MEMORY_ID, key: 'settings', value: { deepseek_model: DEEPSEEK_MODEL }, updated_at: now },
    ], { onConflict: 'pin_hash,key' });
  } catch (e) {
    console.warn('Memory save failed:', e);
  }
}

function scheduleMemorySave() {
  if (!USER_AUTH_HASH) return;
  if (memorySaveTimeout) clearTimeout(memorySaveTimeout);
  memorySaveTimeout = setTimeout(() => saveMemoryToCloud(), 2000);
}

// ── INIT ──────────────────────────────────────────────────────────────────────
buildSidebar();

// Load from Supabase cloud database
(async () => {
  // Auto-login from localStorage if auth was saved
  try {
    const savedHash = localStorage.getItem('sfa_auth_hash');
    if (savedHash) {
      USER_AUTH_HASH = savedHash;
      updateMemoryBadge();
      loadMemoryFromCloud().catch(() => {});
    }
  } catch(e) {}

  if (SUPABASE_URL && SUPABASE_ANON_KEY) {
    const loaded = await loadFromSupabase().catch(e => { console.error('Init error:', e); return false; });
    if (loaded) {
      // Show login modal on first visit (not signed in)
      if (!USER_AUTH_HASH) {
        setTimeout(() => showPinModal(), 1500);
      }
      return;
    }
  }
  // Fallback: try IndexedDB cache
  const cached = await tryLoadFromCache().catch(() => false);
  if (!cached) {
    document.getElementById('hdrSearch').placeholder = lang === 'ar' ? 'لا توجد بيانات' : 'No data available';
  }
})();
