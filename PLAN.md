# US M&A Tax — Resource Scheduling & Capacity Planning: Design Plan

**Prepared for:** Leo Berwick US M&A Tax team (per Justine Morin's brief of June 4)
**Inputs reviewed:** "Scheduling M&A US Tax" email thread · `Scheduling_US_MA_July_9.xlsx` (current staffing tracker)
**Status:** Draft for review. Design decisions locked so far are in §3; open discovery questions in §8.

---

## 1. The ask (from the email)

Design and implement a scheduling and resource-allocation solution for the US M&A Tax team that:

- Improves resource allocation across transactions and balances workload across the team
- Supports professional development (exposure across deal types and workstreams)
- Provides forward visibility on pipeline and capacity
- Is scalable as the team grows and embeds into existing systems (NetSuite, internal workflows)

The core loop Justine describes:

1. A deal enters the pipeline
2. The system assigns an **expected effort profile** (estimated hours by level, Associate → Partner — directional, not precise)
3. Effort translates into a **staffing mix across levels** (Associate-heavy vs VP-heavy, etc.)
4. A **forward schedule** is generated across the team (time-phased: diligence spike, quieter post-signing)
5. The schedule **updates dynamically** as deals evolve (timing, probability, scope, reassignment)

Supporting requirements: weekly capacity by resource with over-allocation / under-utilization flags; scenario analysis; a weekly 15–20 min governance check-in for VP+ to validate near-term capacity and confirm staffing; a NetSuite actuals-vs-estimates feedback loop; rules-based estimation first, AI-enhanced over time.

## 2. Current state — what the July 9 workbook shows

### 2.1 Contents

- **`Sheet1` (deal register):** 292 coded engagements — 195 WIP, 90 Proposal, plus a handful of Dead/Closed/other. Columns: Client, Project ID, WIP-vs-Proposal, then one staffing column per level (Partners, MD, Director, VP, SA, Renewable Specialist) holding free-text first names. A "Status – Complete and ready to invoice" column exists but is entirely empty.
- **Roster (embedded in side columns):** ~23 people — 4 Partners, 4 MDs, 4 Directors, 3 VPs, 4 Senior Associates, and a 4-person Renewable specialist pool.
- **`Summary by Person`:** a manually assembled per-person survey of engagement statuses, in two different layouts depending on who filled it in, with **34+ distinct free-text status values** (including three spellings of "Pencils Down" and entries like "Don't know what this is").
- **`Filtered Status`:** a manual cleanup pivot for a few people — evidence someone is already spending time reconciling this by hand.

### 2.2 Gaps between current state and the target

| # | Gap | Consequence today | What the design must add |
|---|-----|-------------------|--------------------------|
| 1 | **No time dimension** — no dates, hours, or % anywhere | Capacity math is impossible; the sheet says *who* is on a deal, never *how much* or *when* | Planned hours per person-deal-week; deal dates (expected start / sign / close) |
| 2 | **No deal attributes** | The estimation drivers from the email (transaction type, entity class, scope, industry, complexity, timeline) aren't captured | Structured attribute fields at intake, driving effort templates |
| 3 | **Identity ambiguity** — free-text first names | "Kyle" is a Director on some rows and a Partner on others (two Kyles); four variants of Will; case/whitespace variants; multi-name cells ("Tania / Will", "Dorian, Yiyi, Robbie") | Canonical person records with IDs; migration disambiguates via the level column a name appears in, with a human review list for leftovers |
| 4 | **Status chaos** — 34+ hand-typed values mixing lifecycle and phase | Nobody can filter reliably; the invoicing column was abandoned | Controlled lifecycle status + separate deal-phase field (§4.4) |
| 5 | **No probability on the 90 proposals** | Pipeline-weighted capacity is impossible; proposals are a third of the register | Probability field (bucketed), set at the weekly check-in |
| 6 | **Hierarchy mismatch** | Email lists Associate → Partner (6 levels); the sheet has no Associate column, and "Renewable Specialist" is a cross-cutting pool the email doesn't mention | Level *and* specialty modeled separately; confirm Associate plans in discovery |
| 7 | **Referral chains in the client field** ("A : B : C") | Referral source, paying entity, and end client are conflated | Client entity with role (referral source / paying entity / end client) |
| 8 | **Volume skew** | Renewable specialists carry very high engagement counts (one person appears on 75 projects) — many small reviews alongside chunky M&A deals | Effort templates must cover both shapes: high-volume/low-hours advisory and full-scope deals |

**Scale reality check:** ~23 people × ~300 engagements × 52 weeks is tiny data. Nothing here is a performance problem — the hard parts are workflow design, estimation quality, and adoption. "Scalable" means organizationally scalable (new levels, new people, sub-teams), which is a data-model concern, not an infrastructure one.

## 3. Design decisions locked so far

| Decision | Choice | Implication |
|----------|--------|-------------|
| Delivery form | **Hybrid path** | Ship a cleaned, structured workbook in week 1–2 for immediate relief; build the web app in parallel; the workbook doubles as the schema prototype and migration source |
| MVP scope | **Thin end-to-end slice** | First app release covers the whole loop shallowly: tracker + simple hours-by-level estimates + person×week capacity heatmap with flags — then each part deepens |
| Effort & capacity unit | **Hours per week** | Matches NetSuite time entries, so estimate-vs-actual calibration is direct. Entry stays template-driven — nobody hand-types hour grids |
| NetSuite | **CSV export first** | A saved-search export of time entries by project code, imported weekly. Feedback loop from day one without an API project; automate via API in Phase 4 |

## 4. Target concept

### 4.1 Data model (core entities)

- **Person** — canonical name, level (Associate → Partner), specialty tags (e.g., Renewable), weekly capacity hours, active dates, development goals (target exposure by deal type / workstream)
- **Client** — canonical name, with roles per engagement: referral source vs paying entity vs end client
- **Deal** — project code (matching NetSuite), client links, lifecycle status, **probability %**, key dates (expected start / sign / close), attributes: transaction type (stock / asset / partnership), entity classification (corp / flow-through), scope set (TDD, structuring, modeling, PW&A, …), industry, complexity flags (tax equity, cross-border, carve-out, …), timeline (compressed / standard)
- **Effort template** — attribute pattern → baseline hours by level + a phase curve (how those hours spread over the deal's weeks)
- **Assignment** — person × deal × role level → planned hours per week (generated from the template's staffing mix, then editable)
- **Actuals entry** — NetSuite time rows (person, project code, week, hours)
- **Scenario** — an overlay of changed dates / probabilities / assignments for what-if analysis, never touching the live plan until applied

### 4.2 The scheduling loop (engine)

1. **Intake:** deal created with attributes → matched to an effort template → estimated hours by level (directional, overridable)
2. **Phasing:** hours spread across the timeline via the phase curve (diligence spike → pre-close → post-close taper)
3. **Staffing:** system suggests individuals per level based on available capacity, level fit, and development goals; a senior confirms
4. **Ledger:** person × week planned hours (probability-weighted and unweighted views) vs capacity → utilization %, with over-allocation and under-utilization flags (thresholds set in discovery)
5. **Update loop:** status / date / probability changes reflow *future* weeks only; past weeks stay as record
6. **Calibration:** weekly NetSuite CSV → actual vs estimate by deal and level → suggested template multiplier updates (rules-based; a fit/AI layer once enough history accumulates)

### 4.3 Views

- **Pipeline board** — live + expected deals, filterable by status, probability, attributes
- **Capacity heatmap** — person × week, color-coded utilization, the team's main screen
- **Deal page** — staffing, estimate vs actual burn, dates, history
- **Person page** — load ahead, current mix of deal types / workstreams vs development goals
- **Check-in view** — auto-generated agenda for the weekly 15–20 min: people over-allocated in the next 4 weeks, unstaffed or under-staffed incoming deals, aging proposals, stale statuses
- **Scenario compare** — e.g., "these two deals close simultaneously" or "deal slips 3 weeks" side-by-side with the live plan

### 4.4 Status taxonomy (proposed — replaces 34 free-text values)

Two separate fields:

- **Lifecycle:** `Proposal` → `Active` → (`On hold` | `Pens down`) → `Complete – to invoice` → `Invoiced` → `Closed`, with `Dead` reachable from Proposal/Active
- **Phase** (active deals only): `Pre-LOI` · `Diligence` · `Signing → Close` · `Post-close`

The current sheet mixes these into one column (e.g., "TDD done / Structuring Ongoing", "Invoicing soon") — splitting them is what makes both the workload curve and the invoicing pipeline reportable.

## 5. Delivery plan — hybrid, two tracks

### Track A — structured workbook (weeks 1–2, immediate relief)

1. **Cleaning/migration script** (Python) run against the July 9 file:
   - Normalize names against the canonical roster; disambiguate via the level column each name appears in; split multi-name cells; fix case/whitespace
   - Parse client chains into referral source / paying entity / end client
   - Map the 34 status values onto the §4.4 taxonomy
   - Emit a short **review list** of genuinely ambiguous rows for a human to confirm
2. **Workbook v2**, generated by the script:
   - `Roster` — canonical people, level, specialty, weekly capacity hours
   - `Deals` — one row per engagement: attributes, probability, expected dates, lifecycle + phase via dropdown validation
   - `Assignments` — one row per person-deal, with estimated hours by level
   - `Capacity` — person × week pivot of planned hours with conditional-formatting flags
   - Legend naming the cells to edit + one example row
3. This workbook immediately serves the weekly check-in, and is the vehicle for collecting the missing data (dates, probabilities, attributes) that the app will import. It is retired as soon as the MVP is adopted — it is a bridge, not a second system to maintain long-term.

### Track B — web app (parallel)

- **Phase 1 — MVP, thin end-to-end slice:** sign-in (likely Microsoft 365 SSO — the firm is on Outlook/Exchange), CRUD for roster/deals/assignments, template-picker effort estimates with override, person×week capacity heatmap with flags, CSV import (workbook v2 + NetSuite time export), audit trail of changes
- **Phase 2 — estimation depth:** template library workshopped with VP/Director/MD to encode current heuristics per deal archetype; phase curves; probability-weighted pipeline view
- **Phase 3 — dynamics:** scenario sandbox, reassignment workflow, auto-generated check-in agenda, notifications (e.g., "you were added to X", "Y is over 100% in two weeks")
- **Phase 4 — integration & learning:** NetSuite API replaces the CSV; calibration dashboard (estimate vs actual by template and level) with suggested multiplier updates; AI-enhanced estimation once history accumulates; growth features (new levels, sub-teams)

**Suggested stack (to confirm):** TypeScript + React front end, Node or Python API, Postgres (SQLite is honestly sufficient at this scale to start). Private hosting with role-based access — client names and deal codenames are confidential M&A information.

## 6. Data hygiene & migration

The July 9 file is the seed data. The migration script (Track A step 1) is written once and kept: it becomes the importer the app uses, so the cleanup effort is not throwaway. Expected human touchpoints: the ambiguous-name review list and sign-off on the status mapping — likely a 30-minute review, not a re-keying exercise.

## 7. NetSuite feedback loop (CSV-first)

1. NetSuite admin builds a saved search: time entries grouped by project code × person × week
2. Weekly export (CSV) dropped into the tool → matched on project code (the register's "Project ID" appears NetSuite-shaped; confirm in discovery)
3. Deal pages show estimate vs actual burn; calibration view aggregates variance by template
4. Phase 4 replaces the manual export with SuiteTalk/REST on a schedule — same pipeline, different transport

## 8. Open questions for discovery

1. **Roster ground truth** — are there Associates today or planned hires? Anyone missing from the side-column roster? Who owns roster updates?
2. **Renewable pool** — do the 4 specialists staff only renewable/insurance-review work or blend into M&A deals? They likely need their own effort templates (high-volume, low-hours engagements).
3. **Capacity baseline** — standard weekly hours per level? Billable targets? Where do PTO/holidays come from (HR system, Outlook calendars)?
4. **Probability buckets** — e.g., 25/50/75/90%? Who sets and updates them (proposal owner vs check-in)?
5. **Dates at intake** — what is actually known when a deal enters (expected sign/close)? Typical phase durations by scope?
6. **NetSuite specifics** — do register "Project ID"s match NetSuite project codes exactly? Entry granularity (task-level?)? How much history is available for calibration? Who can build the saved search?
7. **Access & permissions** — who edits vs views? Confirm Microsoft 365 SSO. Any hosting constraints (approved cloud, region)?
8. **Development goals** — how formal should this be (e.g., "each SA sees ≥2 deal types per quarter")? It drives the staffing suggester.
9. **Non-deal time** — BD, proposal writing, internal projects, training: model as pseudo-engagements so capacity is honest?
10. **Confidentiality walls** — any deals restricted to named staff, requiring per-deal access control from day one?

## 9. Immediate next steps

1. Review this plan (Justine / Tomas / Zack) and answer §8 — a single 30-minute session covers most of it
2. Build Track A: cleaning script + workbook v2 from the July 9 file, producing the ambiguity review list
3. Scaffold Track B: repo structure, schema, importer reusing the Track A script
4. Template workshop with VP/Director/MD to draft the first effort templates (can run in parallel with MVP build)
