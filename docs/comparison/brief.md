# Comparison brief

The same request and notes were given to Claude twice: once with Anthropic's `pptx` skill, once with Slidewell. Everything below is fictional.

## Request

> Make a 10-slide deck for our board recommending whether Harbor Roast should open in Northfield. The board meets in February 2027 and will read it before the meeting and then decide. Use our brown and gold brand. Analysis notes below.

## Analysis notes

**Market (specialty coffee spend, £M, 2020–2025 actual, 2026–2030 team forecast)**
- Northfield: 120, 131, 143, 156, 170, 185 → 200, 216, 233, 251, 270 (about 8% a year)
- Home region: 210, 218, 226, 233, 240, 247 → 254, 260, 267, 273, 280 (about 3% a year)
- Source: Regional Consumer Spend Survey 2025; team forecast

**Districts (weekday footfall, thousands a day; specialty cafés per 10,000 residents)**
- University Quarter 48, 0.8 · Riverside Campus 41, 1.1 · Old Town 52, 3.8 · Station 45, 3.1 · Market Hill 35, 2.6 · Eastgate 22, 1.8 · Westbrook 18, 1.4 · Retail Park 30, 1.3
- Students in the two university districts match our home-region core customer
- Source: council footfall counters 2025; café listings January 2027

**Store economics (one store, steady-state year 2 onwards, £K)**
- Revenue 1,000; coffee and food 310; staff 280; rent 170; other 100; operating profit 140 (14% margin vs 11% home average)
- Rent is 17% of revenue, three points below home; staff in line with home at local wage rates
- Fit-out £350K; year 1 profit 112 (80% of steady state); payback after 2 years 8 months
- Source: home-region store P&Ls 2026; Northfield rent quotes

**Options (rated against the board's criteria agreed in November 2026: value, risk, speed, brand control)**
- Six-store pilot: value fair, risk good, speed good, brand control good → overall good
- 20-store rollout: good, poor, fair, good → fair
- Franchise partner: fair, good, good, poor → fair
- Do nothing: poor, good, good, good → poor

**Risks and mitigations**
- Leases taken by competitors (medium): sign heads of terms for four sites by March, two backups identified
- Can't hire experienced managers locally (medium): transfer two home-region managers for year 1, train local deputies
- Student demand drops in summer (high): summer menu and hours, target staff and visitors
- Fit-out costs overrun (low): fixed-price contracts, 10% contingency inside the £2.4M

**Ask:** £2.4M to open six stores between March and September 2027 (6 × £350K fit-out = £2.1M, plus £0.3M launch and contingency). Plan: sign leases January–February; stores 1–2 open March, 3–4 June, 5–6 September; board review December.

## How the two decks were made

Both were built by Claude from this brief. The Slidewell deck is `slidewell/examples/market-entry-recommendation.pptx`, built with Slidewell. The pptx-skill deck was built afterwards (November 2026) following Anthropic's `pptx` skill as written: pptxgenjs, a theme with named layouts and placeholders, its design ideas, type sizes and QA steps, and no Slidewell references. The Claude that built it had seen the Slidewell deck, which if anything helped the pptx-skill version (its headlines are better than a cold start would give). Both were then run through Slidewell's own checker, so treat the checker counts as Slidewell's standard, not a neutral score; the slide images are the fairer comparison. To reproduce, give the request and notes above to Claude with each skill installed in a fresh conversation.

## Results

| | pptx skill | Slidewell |
|---|---|---|
| Slidewell checker | 1 error, 38 warnings | 0 errors, 9 warnings |
| Anthropic's `validate.py` (file validity) | Passed | Passed |

Both decks: `pptx-skill/pptx-skill-deck.pptx` (and its `contact-sheet.png`) and `../../slidewell/examples/market-entry-recommendation.pptx`.
