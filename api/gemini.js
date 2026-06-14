/**
 * Map Peace — Gemini proxy (Vercel / Node serverless).
 * Route: POST /api/gemini   body: { prompt: string, maxTokens?: number }
 * Returns: { text: string }
 *
 * SECURITY: reads process.env.GEMINI_API_KEY (a project secret). Never hardcode it.
 */
const MODEL = "gemini-2.0-flash";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Use POST" });

  const key = process.env.GEMINI_API_KEY;
  if (!key) return res.status(500).json({ error: "Server not configured: set GEMINI_API_KEY." });

  const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
  const prompt = (body.prompt || "").toString().slice(0, 12000);
  if (!prompt) return res.status(400).json({ error: "Missing prompt" });
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
    if (!r.ok) return res.status(502).json({ error: "Gemini error", detail: (await r.text()).slice(0, 400) });
    const data = await r.json();
    return res.status(200).json({ text: data?.candidates?.[0]?.content?.parts?.[0]?.text || "" });
  } catch (e) {
    return res.status(500).json({ error: "Request failed", detail: String(e).slice(0, 300) });
  }
}
