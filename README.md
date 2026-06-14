# Map Peace — A Personal Guide to Al-Madinah Al-Munawwarah

Map Peace helps visitors to Al-Madinah know **where they should go**. A visitor
answers a short questionnaire and instantly gets a personal, day-by-day plan with
**real places**, an **interactive map**, **booking links**, optional **AI
suggestions**, and a **styled PDF** download.

> "Find your peace." — a gentle companion for the city of the Prophet ﷺ.

## What it does

- **Guided questionnaire** (one tap per question): age, month, Ramadan, length of
  stay, company, goal, daily pace, transport, outdoor preference, fitness,
  budget, and Rawdah planning.
- **Personalized itinerary engine** that builds a day-by-day plan from the answers:
  - Always anchors plans on **Al-Masjid an-Nabawi** and **Quba Mosque** (musts).
  - **Tourist / half** trips are attraction-led (the Haram no longer dominates
    every day); **religious** trips stay worship-centred.
  - **Weather-aware** date-farm timing by visit month.
  - **Via Ferrata** is included for young, active, outdoor-loving visitors with
    friends on a half-religious or tourist trip.
  - Real Madinah places: **Al Marbad Farm**, **Al-Aliya date farm**, **King Fahd
    Central Park**, **Uhud E-Bike trail**, **Uhud quad / 4WD**, **Wadi Al-Jinn**
    (Magnetic Valley), **Hijaz Railway** & **Dar Al Madinah** museums, the
    **Qur'an Printing Complex**, the dates market, malls, and an **AlUla** day trip.
- **Interactive map** (Leaflet + OpenStreetMap, no key) pinning every stop, with
  **📍 Open in Maps** and **🎟 Reserve / Info** links per activity (Nusuk, Visit
  Saudi, Hala Yalla, Sightscape, Welcome Saudi).
- **Real photos** for each place (Creative-Commons via LoremFlickr) with a colour
  fallback so nothing ever looks broken.
- **✨ AI "Discover more places"** — an optional button that asks Google **Gemini**
  for fresh, lesser-known spots matching the visitor's trip (see setup below).
- **Output**: shown on the page, **styled PDF** (html2canvas + jsPDF; print
  fallback), and a standalone **.html** download.

## Tech & structure

```
index.html                 # the whole site (HTML + CSS + JS), no build step
functions/api/discover.js  # AI endpoint for EdgeOne Pages / Cloudflare Pages
api/discover.js            # AI endpoint for Vercel / Node serverless
.env.example               # GEMINI_API_KEY placeholder (never commit real keys)
```

The site works fully **without** the AI function — the questionnaire, map,
booking links, photos and PDF are all client-side. The AI button only needs the
serverless function when you want live suggestions.

## Run locally

```bash
python3 -m http.server 8000   # then open http://localhost:8000
```

The `/api/discover` AI button needs a function host (below); locally you can run
it with the Vercel CLI (`vercel dev`) or Wrangler (`wrangler pages dev .`).

## Enabling the AI "Discover more" feature

1. Get a **Gemini API key** from Google AI Studio.
2. Set it as a **secret / environment variable** named `GEMINI_API_KEY` in your
   hosting dashboard — **do not** put it in the code or commit it.
   - **EdgeOne Pages / Cloudflare Pages**: uses `functions/api/discover.js`. Add
     the variable under the project's *Environment variables / Secrets*.
   - **Vercel**: uses `api/discover.js`. Add the variable under
     *Settings → Environment Variables*.
   - **Netlify**: rename the handler to Netlify's signature, or proxy to the
     Vercel/Pages function; set the env var in *Site settings → Environment*.
3. Deploy. The button calls `POST /api/discover` and renders the returned places.

> 🔐 **Security:** never hardcode or commit the key. If a key is ever shared in
> plaintext (chat, screenshot, commit), **rotate/revoke it** and issue a new one.

## Deploy

Static hosting for the site; a Pages/serverless function for the AI endpoint:
**EdgeOne Pages**, **Cloudflare Pages**, or **Vercel** all serve both from this
repo as-is.

## Notes

Plans are friendly suggestions, not official timings. Always confirm prayer
times, Rawdah / Nusuk reservations, booking availability, and opening hours
before visiting. Photos are representative and easy to swap for your own files.
