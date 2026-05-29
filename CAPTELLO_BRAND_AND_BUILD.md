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

## DECK blueprint (customer-facing .pptx) — 16:9

Presentation standards: dark cover → light content → dark close; one idea per slide;
conclusion headlines (≤36 chars, one red keyword max); 5–25 words per slide; 3-second
glance test. Customer-facing only — NEVER put confidential watch-outs, ticket numbers,
contract gaps, or "send as <rep>" notes on a slide.

**Slide count by call type:**
- Returning / already-kicked-off → **5 slides, NO "Your package, at a glance" slide.**
- First onboarding call → **6 slides, WITH the package slide** (purchased SKUs only),
  inserted after the status slide.

1. **Cover (dark):** white-text logo top-left; red eyebrow "ONBOARDING · WORKING SESSION";
   large white client name; one-line subhead; "Prepared by Captello Onboarding &
   Enablement"; red quarter-circle arc bleeding off bottom-right.
2. **Status / recap (light):** conclusion headline; 3 short bullets left; 3-stage vertical
   progress indicator right (Kicked off → Working sessions [NOW, red pill] → Event-ready).
3. **Agenda (light):** "Two Sessions to Launch-Ready" — two session cards, 3 items each,
   from the canonical workflow (Verify Integrations → Build Template → Create Capture Form
   → Invite Staff → Set Up Follow-Up → Establish Workflow → CRM/MAP Data Flow).
4. **Preparation (light):** "Your Prep Drives a Flawless Event" — 3–4 item customer prep
   checklist with red circular check icons + a dark "Test before launch" card on the right.
   Lead with preparation.
5. **Close (dark):** bold CTA headline; "Your onboarding specialist is one message away";
   "captello.com" in red; white-text logo; red arc.
- **(First call only) package slide:** purchased SKUs as a clean icon/label grid.

## BRIEF blueprint (internal .docx) — US Letter

Internal register: direct, instructional. Confidential — carries contract terms and ticket
detail.

- Header: black-text logo + thin gray rule.
- Eyebrow "CONFIDENTIAL · INTERNAL USE ONLY" (red); title + red rule; status line.
- **Account snapshot** + contacts table (contact / role).
- **Watch-outs for <owning specialist>** — bullets.
- **Open Pylon tickets** (returning only) — table (ticket / issue / status), red status
  cells, flag the oldest ticket's age.
- **Pre-call email draft** — To / Cc routing + "review and send as <rep>; draft only."
- **Recommended agenda for the next call** — numbered.
- Document footer (above).

---

## Build mechanics

The run environment starts blank but can install packages. At the start of the build:

```
pip3 install python-pptx python-docx --quiet      # or: npm install pptxgenjs
```

Download the logo files from THIS Drive folder (Drive connector) into the working dir
before embedding. Build with `scripts/build_deck.py` and `scripts/build_brief.py` (in this
folder) — pass the client name, call type (returning|first), and the synthesized content.
Then upload the finished .pptx/.docx back to Drive, set sharing to domain/team reader
(never "anyone with the link"), and post wrapped hyperlinks in the Slack thread.
