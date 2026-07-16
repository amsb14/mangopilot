// Fetch clean article text from a news URL.
// Primary: Jina AI Reader (r.jina.ai) — handles bot-blocking + JS render + parsing.
// Fallback: raw fetch + lightweight HTML extraction.
// Auth: x-auth-hash (same as other functions). Never hangs — hard 9s timeout.

const APP_PASSWORD_HASH = Deno.env.get("APP_PASSWORD_HASH")!;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-auth-hash, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const MAX_CHARS = 5000;
const TIMEOUT_MS = 9000;

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  const authHash = req.headers.get("x-auth-hash");
  if (!authHash || authHash !== APP_PASSWORD_HASH) return json({ error: "Unauthorized" }, 401);

  let url: string;
  try {
    const body = await req.json();
    url = body.url;
    if (!url || typeof url !== "string") throw new Error("Missing url");
    const parsed = new URL(url);
    if (!["http:", "https:"].includes(parsed.protocol)) throw new Error("Bad protocol");
  } catch (e) {
    return json({ error: `Invalid request: ${(e as Error).message}` }, 400);
  }

  // Try Jina Reader first
  const jinaText = await tryJina(url);
  if (jinaText && jinaText.length > 200) {
    return json({ url, text: jinaText.slice(0, MAX_CHARS), source: "jina", length: jinaText.length });
  }

  // Fallback: raw fetch + basic extraction
  const rawText = await tryRawFetch(url);
  if (rawText && rawText.length > 200) {
    return json({ url, text: rawText.slice(0, MAX_CHARS), source: "raw", length: rawText.length });
  }

  return json({ url, text: null, source: "none", error: "Could not extract article text" });
});

async function tryJina(url: string): Promise<string | null> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  try {
    // r.jina.ai returns clean markdown of the article. X-Return-Format: text strips markdown.
    const res = await fetch(`https://r.jina.ai/${url}`, {
      signal: ctrl.signal,
      headers: {
        "Accept": "text/plain",
        "X-Return-Format": "text",
        "X-Timeout": "8",
      },
    });
    clearTimeout(timer);
    if (!res.ok) return null;
    const text = await res.text();
    return cleanText(text);
  } catch {
    clearTimeout(timer);
    return null;
  }
}

async function tryRawFetch(url: string): Promise<string | null> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      signal: ctrl.signal,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
      },
    });
    clearTimeout(timer);
    if (!res.ok) return null;
    const ct = res.headers.get("content-type") || "";
    if (!ct.includes("text/html") && !ct.includes("xhtml")) return null;
    const html = await res.text();
    return extractFromHtml(html);
  } catch {
    clearTimeout(timer);
    return null;
  }
}

function extractFromHtml(html: string): string | null {
  // Strip non-content elements
  let s = html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
    .replace(/<header[\s\S]*?<\/header>/gi, "")
    .replace(/<footer[\s\S]*?<\/footer>/gi, "")
    .replace(/<nav[\s\S]*?<\/nav>/gi, "")
    .replace(/<aside[\s\S]*?<\/aside>/gi, "");

  // Prefer <article> or <main>, else largest content div
  let body = "";
  const art = s.match(/<article\b[^>]*>([\s\S]*?)<\/article>/i);
  if (art && art[1].length > 400) body = art[1];
  else {
    const main = s.match(/<main\b[^>]*>([\s\S]*?)<\/main>/i);
    if (main && main[1].length > 400) body = main[1];
    else body = s;
  }

  const text = cleanText(body.replace(/<[^>]+>/g, " "));
  return text.length > 200 ? text : null;
}

function cleanText(text: string): string {
  return text
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&rsquo;/g, "'").replace(/&ldquo;/g, '"').replace(/&rdquo;/g, '"')
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCharCode(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)))
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}
