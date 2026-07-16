// Seeds ticker_relationships table with AI-generated relationships for one ticker.
// Client calls this once per ticker (in a loop) so the user gets a progress bar.

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") || "";
const APP_PASSWORD_HASH = Deno.env.get("APP_PASSWORD_HASH")!;
const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-auth-hash, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const VALID_TYPES = ["competitor", "supplier", "customer", "partner", "co_dependent", "thematic_peer"];

interface ReqBody {
  ticker: string;
  name?: string;
  sector?: string;
  universe: string[];
  model?: string;
}

interface AIRelationship {
  related_ticker: string;
  type: string;
  evidence?: string;
  strength?: number;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders });
  if (req.method !== "POST") {
    return jsonRes({ error: "Method not allowed" }, 405);
  }

  const authHash = req.headers.get("x-auth-hash");
  if (!authHash || authHash !== APP_PASSWORD_HASH) {
    return jsonRes({ error: "Unauthorized" }, 401);
  }
  if (!OPENAI_API_KEY) return jsonRes({ error: "OpenAI not configured" }, 500);

  let body: ReqBody;
  try { body = await req.json(); } catch { return jsonRes({ error: "Bad JSON" }, 400); }

  const { ticker, name, sector, universe, model } = body;
  if (!ticker || !Array.isArray(universe) || universe.length === 0) {
    return jsonRes({ error: "ticker and universe[] required" }, 400);
  }

  const universeSet = new Set(universe.map((t) => t.toUpperCase()));
  const tickerUC = ticker.toUpperCase();

  // Ask GPT for relationships
  const sys = `You map company business relationships for a financial analysis tool.

For the given ticker, return up to 8 of its most important business relationships with OTHER PUBLIC US-listed companies.

Relationship types:
- competitor: competes for the same customers/market (e.g. NVDA <-> AMD in GPUs)
- supplier: sells to the source company (e.g. TSM is a supplier to NVDA)
- customer: buys from the source company (e.g. META is a customer of NVDA's chips)
- partner: strategic partnership/integration (e.g. MSFT <-> OpenAI, AAPL <-> GOOGL search deal)
- co_dependent: deeply intertwined success (e.g. AAPL & TSMC for advanced chips)
- thematic_peer: rides the SAME secular trend without being a direct competitor/supplier/customer.
  Examples: NVDA <-> VRT (both ride AI capex; VRT supplies cooling to AI data centers),
  NVDA <-> BE/CEG (data center power demand from AI training),
  TSLA <-> ALB (EV adoption drives lithium demand),
  ASML <-> SMCI (semiconductor capex cycle).
  Use this when companies will move TOGETHER on the same news (AI boom, EV adoption, cloud capex)
  but neither directly buys from nor competes with the other.

Rules:
- Only include relationships you're confident about based on widely-reported business facts.
- Use US ticker symbols only (NYSE/NASDAQ). No foreign listings, no private companies.
- evidence: ONE short factual sentence describing the relationship.
- strength: 1-10 (10 = critical/existential dependency, 5 = meaningful, 1 = minor).
- DO NOT invent relationships. If you don't know 8, return fewer.
- Output ONLY valid JSON, no markdown, no commentary.

Format:
{"relationships":[
  {"related_ticker":"AMD","type":"competitor","evidence":"Both compete in discrete GPU and AI accelerator markets.","strength":9},
  {"related_ticker":"TSM","type":"supplier","evidence":"TSMC manufactures NVIDIA's GPUs on advanced nodes.","strength":10}
]}`;

  const userMsg = `Source ticker: ${tickerUC}
${name ? `Company name: ${name}` : ""}
${sector ? `Sector: ${sector}` : ""}

Return JSON of relationships only.`;

  const aiRes = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: model || "gpt-4o-mini",
      messages: [
        { role: "system", content: sys },
        { role: "user", content: userMsg },
      ],
      temperature: 0,
      max_tokens: 800,
      response_format: { type: "json_object" },
    }),
  });

  if (!aiRes.ok) {
    const errTxt = await aiRes.text();
    return jsonRes({ error: `OpenAI: ${aiRes.status}`, body: errTxt.slice(0, 300) }, 500);
  }
  const aiData = await aiRes.json();
  const content = aiData.choices?.[0]?.message?.content || "{}";

  let parsed: { relationships?: AIRelationship[] };
  try { parsed = JSON.parse(content); }
  catch { return jsonRes({ error: "AI response not valid JSON", raw: content.slice(0, 200) }, 500); }

  const rels = (parsed.relationships || [])
    .filter((r) => r && r.related_ticker && r.type)
    .map((r) => ({
      source_ticker: tickerUC,
      related_ticker: r.related_ticker.toUpperCase().trim(),
      type: r.type.toLowerCase().trim(),
      evidence: (r.evidence || "").slice(0, 500),
      strength: clamp(parseInt(String(r.strength || 5)) || 5, 1, 10),
    }))
    .filter((r) =>
      r.related_ticker !== tickerUC &&
      universeSet.has(r.related_ticker) &&
      VALID_TYPES.includes(r.type)
    );

  if (rels.length === 0) {
    return jsonRes({ ticker: tickerUC, inserted: 0, raw_count: parsed.relationships?.length || 0 });
  }

  // Upsert into Supabase using service role (bypasses RLS).
  // on_conflict tells PostgREST which columns identify a duplicate so merge-duplicates can fire.
  const dbRes = await fetch(
    `${SUPABASE_URL}/rest/v1/ticker_relationships?on_conflict=source_ticker,related_ticker,type`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SERVICE_ROLE_KEY,
        "Authorization": `Bearer ${SERVICE_ROLE_KEY}`,
        "Prefer": "resolution=merge-duplicates",
      },
      body: JSON.stringify(rels),
    }
  );

  if (!dbRes.ok) {
    const dbErr = await dbRes.text();
    return jsonRes({ error: `DB write failed: ${dbRes.status}`, detail: dbErr.slice(0, 300) }, 500);
  }

  return jsonRes({ ticker: tickerUC, inserted: rels.length, relationships: rels });
});

function jsonRes(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n));
}
