// Server-side proxy for read-only FMP market-data endpoints. The FMP key never
// reaches the browser — same auth pattern as rapid-worker (x-auth-hash).
const FMP_API_KEY = Deno.env.get("FMP_API_KEY") || "";
const APP_PASSWORD_HASH = Deno.env.get("APP_PASSWORD_HASH")!;
const FMP_BASE = "https://financialmodelingprep.com/stable";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-auth-hash, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

// Whitelist of endpoints the client is allowed to reach — read-only market data only.
const ALLOWED_ENDPOINTS = new Set([
  "quote",
  "historical-price-eod/light",
  "historical-chart/5min",
  "historical-chart/1hour",
  "news/stock",
  "ratings-snapshot",
  "price-target-consensus",
  "earnings-calendar",
  "biggest-gainers",
  "biggest-losers",
  "most-active",
]);

// Only pass through scalar (string/number) params, and never let the client override apikey.
function sanitizeParams(params: unknown): Record<string, string> {
  const out: Record<string, string> = {};
  if (!params || typeof params !== "object") return out;
  for (const [k, v] of Object.entries(params as Record<string, unknown>)) {
    if (k === "apikey") continue;
    if (typeof v === "string" || typeof v === "number") out[k] = String(v);
  }
  return out;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: corsHeaders });
  }
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const authHash = req.headers.get("x-auth-hash");
  if (!authHash) {
    return new Response(JSON.stringify({ error: "Authentication required. Please sign in." }), {
      status: 401,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  if (!APP_PASSWORD_HASH || authHash !== APP_PASSWORD_HASH) {
    return new Response(JSON.stringify({ error: "Invalid password. Please try again." }), {
      status: 403,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  if (!FMP_API_KEY) {
    return new Response(JSON.stringify({ error: "Market data not configured on server." }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  let body: { endpoint?: string; params?: unknown };
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Bad JSON" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const endpoint = body.endpoint || "";
  if (!ALLOWED_ENDPOINTS.has(endpoint)) {
    return new Response(JSON.stringify({ error: "Endpoint not allowed" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const qs = new URLSearchParams(sanitizeParams(body.params));
  qs.set("apikey", FMP_API_KEY);

  const upstreamRes = await fetch(`${FMP_BASE}/${endpoint}?${qs}`);
  const text = await upstreamRes.text();
  return new Response(text, {
    status: upstreamRes.status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
