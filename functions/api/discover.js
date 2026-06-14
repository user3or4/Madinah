/**
 * Map Peace — "Discover more places" AI endpoint
 * Pages Function (works on Tencent EdgeOne Pages & Cloudflare Pages).
 * Route: POST /api/discover
 *
 * SECURITY: the Gemini API key is read from the GEMINI_API_KEY environment
 * variable / secret configured in your hosting dashboard. NEVER hardcode it
 * here or commit it to the repository.
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
  if (!key) {
    return new Response(
      JSON.stringify({ error: "Server not configured: set the GEMINI_API_KEY secret." }),
      { status: 500, headers: cors }
    );
  }

  let body = {};
  try { body = await request.json(); } catch (_) {}
  const a = body.answers || {};
  const existing = Array.isArray(body.existing) ? body.existing : [];
  const ctx = body.context || {};

  const prompt = [
    "You are a friendly local guide to Al-Madinah Al-Munawwarah, Saudi Arabia.",
    "Suggest 4 real, currently-open, lesser-known places or experiences in or near Madinah",
    "that suit this visitor. Favour authentic, recently-popular spots (cafes, date farms,",
    "viewpoints, heritage corners, family or adventure spots). Be respectful of the city's",
    "spiritual character. Do NOT repeat any place already in their plan.",
    "",
    "Visitor profile (JSON): " + JSON.stringify(a),
    "Trip context: " + JSON.stringify(ctx),
    "Already in their plan (avoid these): " + JSON.stringify(existing),
    "",
    "Return ONLY JSON of the form:",
    '{"places":[{"name":"","area":"","category":"","why":"","mapsQuery":""}]}',
    "where category is one of: Cafe, Farm, Nature, Heritage, Family, Adventure, Food, Shopping.",
    "mapsQuery should be a precise Google Maps search string including 'Madinah'.",
  ].join("\n");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent`;

  try {
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.9, responseMimeType: "application/json" },
      }),
    });

    if (!r.ok) {
      const t = await r.text();
      return new Response(JSON.stringify({ error: "Gemini error", detail: t.slice(0, 400) }), { status: 502, headers: cors });
    }

    const data = await r.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
    let parsed = {};
    try { parsed = JSON.parse(text); }
    catch (_) {
      const m = text.match(/\{[\s\S]*\}/);
      parsed = m ? JSON.parse(m[0]) : { places: [] };
    }
    const places = Array.isArray(parsed.places) ? parsed.places.slice(0, 6) : [];
    return new Response(JSON.stringify({ places }), { status: 200, headers: cors });
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
