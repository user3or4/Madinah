/**
 * Map Peace — Gemini proxy (Pages Function: EdgeOne Pages / Cloudflare Pages).
 * Route: POST /api/gemini   body: { prompt: string, maxTokens?: number }
 * Returns: { text: string }   (the model's raw JSON text)
 *
 * SECURITY: reads GEMINI_API_KEY from the host's environment/secret store.
 * NEVER hardcode or commit the key.
 */
const MODEL = "gemini-2.0-flash";

export async function onRequestPost(context) {
  const { request, env } = context;
  const cors = {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };

  const key = env && env.GEMINI_API_KEY;
  if (!key) return new Response(JSON.stringify({ error: "Server not configured: set the GEMINI_API_KEY secret." }), { status: 500, headers: cors });

  let body = {};
  try { body = await request.json(); } catch (_) {}
  const prompt = (body.prompt || "").toString().slice(0, 12000);
  if (!prompt) return new Response(JSON.stringify({ error: "Missing prompt" }), { status: 400, headers: cors });
  const maxTokens = Math.min(Math.max(parseInt(body.maxTokens, 10) || 8192, 256), 8192);

  try {
    const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.85, responseMimeType: "application/json", maxOutputTokens: maxTokens },
      }),
    });
    if (!r.ok) return new Response(JSON.stringify({ error: "Gemini error", detail: (await r.text()).slice(0, 400) }), { status: 502, headers: cors });
    const data = await r.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    return new Response(JSON.stringify({ text }), { status: 200, headers: cors });
  } catch (e) {
    return new Response(JSON.stringify({ error: "Request failed", detail: String(e).slice(0, 300) }), { status: 500, headers: cors });
  }
}

export async function onRequestOptions() {
  return new Response(null, {
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
  });
}
