# Map Peace — A Personal Guide to Al-Madinah Al-Munawwarah

Map Peace helps visitors to Al-Madinah know **where they should go**. A visitor
answers a short, friendly questionnaire and instantly receives a personal,
day-by-day plan they can read on screen, **print / save as PDF**, or download.

> "Find your peace." — a gentle companion for the city of the Prophet ﷺ.

## What it does

- **Guided questionnaire** (one tap per question, with progress bar):
  age, month of visit, Ramadan, length of stay, travel company, goal of the
  visit, daily pace, transport, love of the outdoors, fitness, comfort/budget,
  and Rawdah planning.
- **Personalized itinerary engine** that builds the plan from the answers:
  - **Always** anchors every plan on **Al-Masjid an-Nabawi** (the Prophet's
    Mosque ﷺ) and **Quba Mosque** as musts.
  - **Weather-aware** — reads the visit month and recommends a **date-farm**
    experience suited to the season (cool-day vs. evening / premium private farm).
  - **Via Ferrata** mountain climb is included for the right profile — young,
    active, outdoor-loving visitors travelling with friends on a half-religious
    or tourist trip (per the brief).
  - Adds Uhud, Baqi, Qiblatain & the Seven Mosques, the Qur'an Printing Complex,
    Wadi Al-Jinn, hiking, cycling, museums, AlUla day trips, malls, and a
    Ramadan night programme — selected to fit the visitor.
  - Paces the plan by length of stay and how much of the day to fill
    (nights only / half day / full day).
- **Output**: shown on the page, **Save / Print as PDF** (print-optimized),
  and **Download plan (.html)**.

## Design

Calm, spiritual aesthetic — deep emerald green, warm sand/cream, and gold,
with Cormorant Garamond display type and subtle Islamic geometric patterns.
English-first (LTR) for local Saudi visitors now, with foreign visitors in mind.

## Tech

A single self-contained `index.html` (HTML + CSS + JS, no build step, no
dependencies beyond Google Fonts). Deploy anywhere static — EdgeOne, Netlify,
GitHub Pages, etc.

## Run / deploy

Just open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server 8000   # then visit http://localhost:8000
```

## Notes

Plans are friendly suggestions, not official timings. Always confirm prayer
times, Rawdah / Nusuk reservations, and opening hours before visiting.
