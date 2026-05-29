# Captello Brand & Build Spec — Onboarding Prep (read this before building)

This folder is the single source of truth for branding the onboarding deck and brief.
Read this file, then use the logos in `logos/` and the scripts in `scripts/`. Do not
recreate, trace, or approximate the Captello logo — embed the provided files only.

---

## Non-negotiables

- **Colors (only these three):** red `#FF0000`, black `#000000`, white `#FFFFFF`.
- **Font:** Roboto.
- **Logo aspect ratio:** 5.415 (width ÷ height). Embed with ONE dimension set, or match
  this ratio exactly — never stretch, never add effects.
- **Logo selection:** white-text logo on dark backgrounds, black-text logo on light
  backgrounds. Files in `logos/`:
  - `captello-horizontal-blacktext-transparent.png` (light slides / doc header)
  - `captello-horizontal-whitetext-transparent.png` (dark slides)
  - `captello-icon.svg`, plus SVG wordmarks for HTML use.
- **Capitalization:** Captello (capital C), IntelliScan™, captello.com (no www.),
  Connexions (with an X), Lead Liaison (parent). Never "Captello app," "GoCapture!,"
  "lead capture tool," or "Connections."
- **Voice:** bold, confident, direct, professional. Avoid "just," "simple," "basic,"
  "easy," "stuff," "things," and casual tone.

---

## Footers

- **Document footer (brief .docx), ALL CAPS, pipe-separated:**
  `© COPYRIGHT | LEAD LIAISON | CAPTELLO | ALL RIGHTS RESERVED | 13101 PRESTON ROAD STE 110 - 159 DALLAS, TX 75240 | 888.399.6430 | +44 20 38074910`
- **Deck content slides:** a small gray confidentiality notice, right-aligned, next to the
  black-text logo bottom-left.
- **Operational email footer:**
  `CAPTELLO | 13101 PRESTON RD. STE 110 – 159 DALLAS, TX 75240 | 888.399.6430 | CAPTELLO.COM`

---

## DECK blueprint (customer-facing .pptx) — 16:9 — 6 slides for EVERY client (hard cap 7)

Built by `scripts/build_deck.py` from a content JSON. Standards: dark cover -> light
content -> dark close; one idea per slide; conclusion headlines (one red keyword max);
3-second glance test. NEVER put confidential watch-outs, ticket numbers, contract gaps,
or "send as <rep>" notes on a slide.

1. **Cover (dark):** white-text logo, red eyebrow, client name, subhead, prepared-by line, red arc.
2. **Recap / Goals (REQUIRED):** returning -> "Where We Left Off" with a recap of the last
   meeting; first call -> "Your Goals, Our Plan" with the client's goals. Left = bullets,
   right = 3-stage progress tracker (NOW = Working sessions for returning, Kickoff for first).
3. **Your Package, At a Glance (REQUIRED, ALL clients):** purchased SKUs as a card grid.
4. **Two Sessions to Launch-Ready:** two session cards from the canonical workflow.
5. **Your Prep Drives a Flawless Event:** prep checklist + dark "Test before launch" card.
6. **Close (dark):** CTA, "captello.com" in red, white-text logo, red arc.

Every deck is 6 slides. Do not exceed 7.

## BRIEF blueprint (internal .docx) — US Letter

Internal register: direct, instructional. Confidential — carries contract terms and ticket
detail.

- Header: black-text logo + thin gray rule.
- Eyebrow "CONFIDENTIAL · INTERNAL USE ONLY" (red); title + red rule; status line.
- **Account snapshot** + contacts table (contact / role).
- **Purchased products & services** — bulleted list from the contract (REQUIRED).
- **Watch-outs for <owning specialist>** — bullets.
- **Open Pylon tickets** (returning only) — table (ticket / issue / status), red status
  cells, flag the oldest ticket's age.
- **Pre-call email draft** — To / Cc routing + "review and send as <rep>; draft only."
- **Recommended agenda for the next call** — numbered.
- Document footer (above).

---

## ONE-PAGER (rep-facing .html) — built by `scripts/build_onepager.py`

Large fonts, brand colors only, Roboto, inlined SVG logo. Includes: a horizontal 3-stage
progress tracker, a "Your Package" chip row (REQUIRED), recap/goals + challenges, the
two-session plan, key contacts, numbered next steps, and colored open-ticket chips.
Self-contained HTML (logo inlined) — uploads to Drive and renders in-browser; no render
step needed at runtime.

## Build mechanics

The run environment starts blank but can install packages. At the start of the build:

```
pip3 install python-pptx python-docx --quiet
```

Download the logo files from THIS Drive folder (Drive connector) into the working dir
before embedding. Build with `scripts/build_deck.py`, `scripts/build_brief.py`, and
`scripts/build_onepager.py` — pass a content JSON (schemas are in each script header).
Then upload all three into Drive at: <parent shared folder> / <Client Name> / "Onboarding
Pack - <Day, Month D YYYY>". Files inherit the parent folder's team sharing; the Drive MCP
cannot set per-file permissions, so rely on folder inheritance. Post wrapped hyperlinks
(<url|label>) in the Slack thread.
