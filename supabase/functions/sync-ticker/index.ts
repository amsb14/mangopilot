// Sync a small batch of NEW tickers from FMP → Supabase.
// Called from the main app when the user requests adding a missing ticker.
// Uses x-auth-hash (same as deepseek-proxy) — service_role key never leaves the server.

const APP_PASSWORD_HASH = Deno.env.get("APP_PASSWORD_HASH")!;
const FMP_API_KEY = Deno.env.get("FMP_API_KEY") || "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-auth-hash, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const FMP = "https://financialmodelingprep.com/stable";

interface ReqBody {
  tickers: string[];        // ticker symbols to sync (max 10 per call)
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  const authHash = req.headers.get("x-auth-hash");
  if (!authHash || authHash !== APP_PASSWORD_HASH) return json({ error: "Unauthorized" }, 401);
  if (!FMP_API_KEY) return json({ error: "FMP_API_KEY not configured" }, 500);

  let body: ReqBody;
  try { body = await req.json(); } catch { return json({ error: "Bad JSON" }, 400); }

  const tickers = (body.tickers || []).map(t => String(t).toUpperCase().trim())
    .filter(t => /^[A-Z][A-Z.-]{0,5}$/.test(t))
    .slice(0, 10); // limit to 10 per call so we fit inside Supabase's 50s edge-function timeout
  if (!tickers.length) return json({ error: "tickers[] required (1-10, valid ticker format)" }, 400);

  const synced: Record<string, unknown>[] = [];
  const errors: { ticker: string; error: string }[] = [];

  // Sync sequentially per ticker — each ticker is ~7 FMP calls + 4 DB upserts (~2-4s)
  for (const t of tickers) {
    try {
      const result = await syncOne(t);
      synced.push(result);
    } catch (e) {
      errors.push({ ticker: t, error: (e as Error).message });
    }
  }

  // Update sync_metadata so the UI shows "last synced" correctly
  if (synced.length > 0) {
    await sbUpsert("sync_metadata", [{
      id: 1,
      last_sync: new Date().toISOString(),
      ticker_count: synced.length,
      source: "ondemand-sync",
    }], "id").catch(() => {});
  }

  return json({ synced, errors, count: synced.length });
});

async function syncOne(symbol: string) {
  // Fetch everything FMP gives us per ticker, in parallel
  const [inc, bal, cf, prof, qI, qC, est] = await Promise.all([
    fmpGet("income-statement", { symbol, period: "annual", limit: 5 }),
    fmpGet("balance-sheet-statement", { symbol, period: "annual", limit: 5 }),
    fmpGet("cash-flow-statement", { symbol, period: "annual", limit: 5 }),
    fmpGet("profile", { symbol }),
    fmpGet("income-statement", { symbol, period: "quarter", limit: 8 }).catch(() => []),
    fmpGet("cash-flow-statement", { symbol, period: "quarter", limit: 8 }).catch(() => []),
    fmpGet("analyst-estimates", { symbol, period: "annual", limit: 5 }).catch(() => []),
  ]);

  // PROFILE
  const p = Array.isArray(prof) ? prof[0] : prof;
  if (!p) throw new Error("No profile from FMP — ticker may not exist");
  const profileRow = {
    symbol,
    company_name: p.companyName || null,
    sector: p.sector || null,
    industry: p.industry || null,
    exchange: p.exchange || null,
    country: p.country || null,
    price: p.price ?? null,
    market_cap: p.marketCap ?? null,
    beta: p.beta ?? null,
    last_dividend: p.lastDividend ?? null,
    range_52w: p.range ?? null,
    ceo: p.ceo || null,
    employees: p.fullTimeEmployees ?? null,
    description: p.description || null,
    image: p.image || null,
    ipo_date: p.ipoDate || null,
    updated_at: new Date().toISOString(),
  };
  await sbUpsert("company_profiles", [profileRow], "symbol");

  // ANNUAL — merge income + balance + cash flow by fiscal year
  const byYear: Record<number, { i?: Record<string, unknown>; b?: Record<string, unknown>; c?: Record<string, unknown> }> = {};
  (inc || []).filter((r: { period?: string }) => r.period === "FY").forEach((r: Record<string, unknown>) => {
    const y = parseInt(String(r.fiscalYear)) || new Date(String(r.date)).getFullYear();
    if (!byYear[y]) byYear[y] = {};
    byYear[y].i = r;
  });
  (bal || []).filter((r: { period?: string }) => r.period === "FY").forEach((r: Record<string, unknown>) => {
    const y = parseInt(String(r.fiscalYear)) || new Date(String(r.date)).getFullYear();
    if (!byYear[y]) byYear[y] = {};
    byYear[y].b = r;
  });
  (cf || []).filter((r: { period?: string }) => r.period === "FY").forEach((r: Record<string, unknown>) => {
    const y = parseInt(String(r.fiscalYear)) || new Date(String(r.date)).getFullYear();
    if (!byYear[y]) byYear[y] = {};
    byYear[y].c = r;
  });

  const annualRows: Record<string, unknown>[] = [];
  for (const yrStr of Object.keys(byYear)) {
    const yr = parseInt(yrStr);
    const i = (byYear[yr].i || {}) as Record<string, unknown>;
    const b = (byYear[yr].b || {}) as Record<string, unknown>;
    const c = (byYear[yr].c || {}) as Record<string, unknown>;
    annualRows.push({
      symbol,
      fiscal_year: yr,
      revenue: i.revenue ?? null,
      cost_of_revenue: i.costOfRevenue ?? null,
      gross_profit: i.grossProfit ?? null,
      operating_income: i.operatingIncome ?? null,
      net_income: i.netIncome ?? null,
      ebitda: i.ebitda ?? null,
      ebit: i.ebit ?? null,
      eps_diluted: i.epsDiluted ?? null,
      interest_expense: i.interestExpense ?? null,
      tax_provision: i.incomeTaxExpense ?? null,
      shares_diluted: i.weightedAverageShsOutDil ?? null,
      research_and_development: i.researchAndDevelopmentExpenses ?? null,
      sga: i.sellingGeneralAndAdministrativeExpenses ?? null,
      total_assets: b.totalAssets ?? null,
      total_liabilities: b.totalLiabilities ?? null,
      equity: b.totalStockholdersEquity ?? b.totalEquity ?? null,
      total_debt: b.totalDebt ?? null,
      cash: b.cashAndCashEquivalents ?? null,
      current_assets: b.totalCurrentAssets ?? null,
      current_liabilities: b.totalCurrentLiabilities ?? null,
      long_term_debt: b.longTermDebt ?? null,
      retained_earnings: b.retainedEarnings ?? null,
      inventory: b.inventory ?? null,
      receivables: b.netReceivables ?? b.accountsReceivables ?? null,
      net_ppe: b.propertyPlantEquipmentNet ?? null,
      operating_cash_flow: c.operatingCashFlow ?? c.netCashProvidedByOperatingActivities ?? null,
      capital_expenditures: c.capitalExpenditure ? -Math.abs(Number(c.capitalExpenditure)) : null,
      free_cash_flow: c.freeCashFlow ?? null,
      investing_cash_flow: c.netCashProvidedByInvestingActivities ?? null,
      financing_cash_flow: c.netCashProvidedByFinancingActivities ?? null,
      depreciation: c.depreciationAndAmortization ?? null,
      stock_based_comp: c.stockBasedCompensation ?? null,
      dividends_paid: c.commonDividendsPaid ?? null,
      share_repurchase: c.commonStockRepurchased ?? null,
      updated_at: new Date().toISOString(),
    });
  }
  if (annualRows.length) await sbUpsert("annual_financials", annualRows, "symbol,fiscal_year");

  // QUARTERLY
  const quarterlyRows: Record<string, unknown>[] = [];
  (qI || []).filter((r: { period?: string }) => r.period !== "FY").forEach((r: Record<string, unknown>) => {
    const x = (qC || []).find((c: Record<string, unknown>) => c.date === r.date) || {} as Record<string, unknown>;
    quarterlyRows.push({
      symbol,
      fiscal_date: r.date,
      period: r.period ?? null,
      revenue: r.revenue ?? null,
      net_income: r.netIncome ?? null,
      operating_income: r.operatingIncome ?? null,
      gross_profit: r.grossProfit ?? null,
      eps_diluted: r.epsDiluted ?? null,
      operating_cash_flow: x.operatingCashFlow ?? x.netCashProvidedByOperatingActivities ?? null,
      free_cash_flow: x.freeCashFlow ?? null,
      updated_at: new Date().toISOString(),
    });
  });
  if (quarterlyRows.length) await sbUpsert("quarterly_financials", quarterlyRows, "symbol,fiscal_date");

  // ANALYST ESTIMATES
  const estimateRows = (est || []).map((e: Record<string, unknown>) => ({
    symbol,
    date: e.date,
    revenue_avg: e.revenueAvg ?? null,
    revenue_low: e.revenueLow ?? null,
    revenue_high: e.revenueHigh ?? null,
    ebitda_avg: e.ebitdaAvg ?? null,
    net_income_avg: e.netIncomeAvg ?? null,
    net_income_low: e.netIncomeLow ?? null,
    net_income_high: e.netIncomeHigh ?? null,
    eps_avg: e.epsAvg ?? null,
    eps_low: e.epsLow ?? null,
    eps_high: e.epsHigh ?? null,
    sga_expense_avg: e.sgaExpenseAvg ?? null,
    num_analysts_revenue: e.numAnalystsRevenue ?? null,
    num_analysts_eps: e.numAnalystsEps ?? null,
    updated_at: new Date().toISOString(),
  }));
  if (estimateRows.length) await sbUpsert("analyst_estimates", estimateRows, "symbol,date");

  // Return shaped data so the client can merge directly into local state without a full reload
  return {
    ticker: symbol,
    profile: profileRow,
    annual: annualRows,
    quarterly: quarterlyRows,
    estimates: estimateRows,
  };
}

async function fmpGet(endpoint: string, params: Record<string, string | number>) {
  const q = new URLSearchParams({ ...Object.fromEntries(Object.entries(params).map(([k, v]) => [k, String(v)])), apikey: FMP_API_KEY }).toString();
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = await fetch(`${FMP}/${endpoint}?${q}`);
    if (res.status === 429) { await sleep(3000 * (attempt + 1)); continue; }
    if (!res.ok) {
      if (attempt < 2) { await sleep(1500); continue; }
      throw new Error(`FMP ${endpoint}: HTTP ${res.status}`);
    }
    return res.json();
  }
  throw new Error(`FMP ${endpoint}: max retries`);
}

async function sbUpsert(table: string, rows: Record<string, unknown>[], onConflict: string) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/${table}?on_conflict=${onConflict}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SERVICE_ROLE_KEY,
        "Authorization": `Bearer ${SERVICE_ROLE_KEY}`,
        "Prefer": "resolution=merge-duplicates",
      },
      body: JSON.stringify(rows),
    }
  );
  if (!res.ok) {
    const errTxt = await res.text();
    throw new Error(`${table} upsert: ${res.status} ${errTxt.slice(0, 200)}`);
  }
}

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }
function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}
