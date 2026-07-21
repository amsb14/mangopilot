const DEEPSEEK_API_KEY = Deno.env.get("DEEPSEEK_API_KEY") || "";
const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") || "";
const APP_PASSWORD_HASH = Deno.env.get("APP_PASSWORD_HASH")!;

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-auth-hash, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function isOpenAIModel(model: string): boolean {
  const m = (model || "").toLowerCase();
  return m.startsWith("gpt-") || m.startsWith("o1") || m.startsWith("o3") || m.startsWith("o4");
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

  const body = await req.json();
  const { messages, model, max_tokens, temperature, stream, tools, tool_choice, response_format } = body;

  if (!messages?.length) {
    return new Response(JSON.stringify({ error: "Messages are required." }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const useOpenAI = isOpenAIModel(model);

  if (useOpenAI && !OPENAI_API_KEY) {
    return new Response(JSON.stringify({ error: "OpenAI not configured on server." }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  if (!useOpenAI && !DEEPSEEK_API_KEY) {
    return new Response(JSON.stringify({ error: "DeepSeek not configured on server." }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const apiUrl = useOpenAI
    ? "https://api.openai.com/v1/chat/completions"
    : "https://api.deepseek.com/chat/completions";
  const apiKey = useOpenAI ? OPENAI_API_KEY : DEEPSEEK_API_KEY;
  const defaultModel = useOpenAI ? "gpt-5.4-nano" : "deepseek-v4-flash";

  const modelName = model || defaultModel;
  const tokenCap = Math.min(max_tokens || 2048, 4096);
  // GPT-5 family and o-series ("reasoning") models on OpenAI's Chat Completions API
  // require `max_completion_tokens` (not `max_tokens`) and only accept the default
  // temperature. gpt-4o(-mini) also accepts `max_completion_tokens`, so this stays
  // backward-compatible; DeepSeek keeps the classic `max_tokens`.
  const mLower = String(modelName).toLowerCase();
  const isRestricted = mLower.startsWith("gpt-5") || mLower.startsWith("o1") ||
    mLower.startsWith("o3") || mLower.startsWith("o4");

  const upstreamBody: Record<string, unknown> = {
    model: modelName,
    messages,
    stream: stream ?? true,
  };
  if (useOpenAI) upstreamBody.max_completion_tokens = tokenCap;
  else upstreamBody.max_tokens = tokenCap;
  if (!isRestricted) upstreamBody.temperature = temperature ?? 0.7;
  if (tools) upstreamBody.tools = tools;
  if (tool_choice) upstreamBody.tool_choice = tool_choice;
  if (response_format) upstreamBody.response_format = response_format;

  const upstreamRes = await fetch(apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`,
    },
    body: JSON.stringify(upstreamBody),
  });

  if (!upstreamRes.ok) {
    const errText = await upstreamRes.text();
    return new Response(errText, {
      status: upstreamRes.status,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  if (stream) {
    return new Response(upstreamRes.body, {
      status: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
      },
    });
  }

  const data = await upstreamRes.json();
  return new Response(JSON.stringify(data), {
    status: 200,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
