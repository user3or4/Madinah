/**
 * Map Peace — "Discover more places" AI endpoint (Vercel / Node serverless).
 * Route: POST /api/discover
 *
 * SECURITY: reads the key from process.env.GEMINI_API_KEY (a project secret).
 * NEVER hardcode the key here or commit it.
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
  const a = body.answers || {};
  const existing = Array.isArray(body.existing) ? body.existing : [];
  const ctx = body.context || {};

  const prompt = [
    "You are a friendly local guide to Al-Madinah Al-Munawwarah, Saudi Arabia.",
    "Suggest 4 real, currently-open, lesser-known places or experiences in or near Madinah",
    "that suit this visitor. Favour authentic, recently-popular spots. Be respectful of the",
    "city's spiritual character. Do NOT repeat any place already in their plan.",
    "Visitor profile: " + JSON.stringify(a),
    "Trip context: " + JSON.stringify(ctx),
    "Already in plan (avoid): " + JSON.stringify(existing),
    'Return ONLY JSON: {"places":[{"name":"","area":"","category":"","why":"","mapsQuery":""}]}',
    "category in: Cafe, Farm, Nature, Heritage, Family, Adventure, Food, Shopping.",
  ].join("\n");

  try {
    const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.9, responseMimeType: "application/json" },
      }),
    });
    if (!r.ok) return res.status(502).json({ error: "Gemini error", detail: (await r.text()).slice(0, 400) });
    const data = await r.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
    let parsed;
    try { parsed = JSON.parse(text); }
    catch (_) { const m = text.match(/\{[\s\S]*\}/); parsed = m ? JSON.parse(m[0]) : { places: [] }; }
    return res.status(200).json({ places: Array.isArray(parsed.places) ? parsed.places.slice(0, 6) : [] });
  } catch (e) {
    return res.status(500).json({ error: "Request failed", detail: String(e).slice(0, 300) });
  }
}
