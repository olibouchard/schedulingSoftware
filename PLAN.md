# US M&A Tax — Scheduling Tracker in Excel: Delivery Plan

**Prepared for:** Leo Berwick US M&A Tax team (per Justine Morin's brief of June 4)
**Inputs reviewed:** "Scheduling M&A US Tax" email thread · `Scheduling_US_MA_July_9.xlsx` (current staffing tracker)
**Scope decision:** the deliverable is an **Excel workbook**, end to end. The web app is deferred (see §10). Each step below is a self-contained work package sized for one Claude Code session, with a recommended model (Opus vs Sonnet), the inputs to hand the session, and acceptance criteria.

---

## 1. The ask (from the email)

A forward-looking scheduling and resource-allocation process: a deal enters the pipeline → gets an expected effort profile (hours by level, Associate → Partner — directional, not precise) → effort becomes a staffing mix across levels → a time-phased forward schedule (diligence spike, quieter post-signing) → weekly capacity by person with over-allocation / under-utilization flags → updated dynamically as deals move, with a NetSuite actuals-vs-estimates feedback loop and a weekly 15–20 min governance check-in.

## 2. Current state — what the July 9 workbook shows

- **292 register rows** (195 WIP, 90 Proposal, rest Dead/Closed/other) across **~23 people**: 4 Partners, 4 MDs, 4 Directors, 3 VPs, 4 Senior Associates + a 4-person Renewable specialist pool.
- **No time dimension anywhere** — no dates, hours, or % → capacity math is impossible today.
- **No deal attributes** (transaction type, scope, complexity…) → the effort-estimation drivers aren't captured.
- **Identity ambiguity** — free-text first names; two Kyles (Kidd = Partner, Risser = Director) and four Will variants, resolvable only by which level-column a name sits in; multi-name cells ("Tania / Will", "Dorian, Yiyi, Robbie").
- **Status chaos** — 34+ hand-typed per-person status strings (three spellings of "Pencils Down") mixing deal lifecycle with personal involvement.
- **No probability** on the 90 proposals; **1 duplicate register row** (`Cosmetic _1`); **22 engagements** exist only on the per-person survey tab, not in the register.
- The "Summary by Person" tab is a manual survey — exactly the coordination overhead to eliminate.

## 3. Decisions locked

| Decision | Choice |
|---|---|
| Platform | **Excel workbook** (web app deferred) |
| First deliverable | Thin end-to-end slice: tracker + hours-by-level estimates + person×week capacity heatmap |
| Effort & capacity unit | **Hours per week** (matches NetSuite; entry is template-driven, nobody hand-types grids) |
| NetSuite actuals | **CSV export** of time entries, pasted/imported into the workbook |
| Probability prefill | WIP = 100%, Proposal = 50% (flagged yellow for review) |
| Blank deal dates | Deal counts in **every** week of the capacity view until dates are filled |

## 4. Workbook design (the product spec)

Seven tabs. This is the contract every later step builds on — don't restructure it casually.

| Tab | Purpose | Key columns / cells |
|---|---|---|
| **Guide** | How to use, color legend, baked-in assumptions | — |
| **Capacity** | Person × week heatmap of planned hours; the dashboard | A person · B capacity · C+ 26 weekly columns; Team planned / capacity / headroom rows below; red > capacity, amber > 85% |
| **Check-in** | Weekly 15–20 min governance agenda (Step 5), formula-only | §1 mirrors Roster row-for-row: Person/Level/Cap/Max load (next 4 wks)/Max utilization/Flag (Over-allocated >100%, Under-utilized <50%). §2 mirrors Deals row-for-row, Active/Proposal only: Project ID/Client/Lifecycle/Probability/#staffed/Dates set?/Added on/Weeks since added/Flag (space-joined: Unstaffed, MissingLevel, DefaultProbability, NoDates, Stale)/**Timeline status (Step 10, mirrors Deals!BQ; row highlighted red when Overdue)**. Reads Settings!B19 (as-of date, update before each meeting) and B20 (stale-weeks threshold) |
| **Scenario** | What-if analysis (Step 7), isolated from Live | §1 per-person impact (Live peak vs Scenario peak over next 4 wks, Change, Scenario util, Flag: NEWLY OVER / over / eased). §2 scenario per-person×week heatmap (mirror of Capacity, computed from the Assignments scenario-phasing columns) + team Live-vs-Scenario weekly totals. Driven entirely by the Deals violet override columns; blank overrides ⇒ Scenario ≡ Live. **Read-only; never affects Capacity/Check-in** |
| **Deals** | One row per engagement | A Project ID · B client as filed · C primary client · D end client · E referral? · F lifecycle · G phase · H probability · **I/J Kick-off date / Delivery date (renamed Step 10, cells/formulas unchanged)** · **K–P retired & hidden (Step 8; were the legacy attribute placeholders, empty for every real deal — not deleted)** · Q # staffed (formula) · R planned hrs/wk (formula) · S source status · T notes · U effort archetype (legacy input) · V/W front/tail intensity (f) · X front fraction (f) · Y phase split date (f) · Z Added on (input) · AA/AB/AC scenario prob/start/end (violet input) · AD–AG effective scenario + split (f) · **AH–AX workstream scoping (Step 8, input dropdowns — Justine's verbatim taxonomy: TDD deal type · Domestic only? · TDD scope · Other TDD scoping · 9 Modeling ticks · 3 Structuring ticks · Legal docs review)** · **AY–BC workstream fees $ (Step 9, input, one per group) · BD fee $ total (f) · BE–BI duration wks/group (f) · BJ–BP fee hrs/wk by level (f) — BE–BP hidden calc helpers** · **BQ Timeline status (f, Step 10 — Tentative/Not started/Kick-off ≤1 wk/In flight/Delivery ≤2 wks/Overdue/Delivered, off Settings!B19 + I/J + F)** |
| **Assignments** | One row per person × deal; the staffing ledger | A person · B roster level (formula) · C project ID · D client (f) · E deal lifecycle (f) · F staffed-as level · G **Active?** · H override hrs/wk · I planned hrs/wk used (f: override → **fee-driven (Step 9)** → archetype rate → level default) · J deal probability (f) · K weighted hrs/wk (f) · L/M deal start/end (f) · N/O effective start/end (f, sentinel dates when L/M blank) · P source status · Q notes · **R–AA live phasing helpers** (archetype, front/tail x, split date, seg1 start/end/hrs, seg2 start/end/hrs — see §7; seg1 window now also honours dates on a fee-priced deal) · **AB–AN scenario mirror** of the phasing block (Step 7: scn prob/weighted/dates/split/seg1/seg2, feeding the Scenario tab) · **AO deal fee total (f, the has-fee gate) · AP fee hrs/wk per person (f, Step 9: level hrs ÷ active assignees at that level)** |
| **Roster** | Team list | A person · B level · C specialty · D weekly capacity hrs · E active assignments (f) · F committed hrs/wk (f) · G utilization (f, CF flags) |
| **Templates** | Legacy archetypes (Step 3) + **fee-allocation parameters (Step 8)** | Rows 2–10: archetype table (now legacy; unchanged, still referenced by the phasing formulas so it must not move). Rows 14+: **fee→hours parameters** — rate card $/hr by level (17–23) · fee split-% matrix, workstream group × level with a =100% checksum + red CF (26–31) · default duration per group (34–39). All yellow placeholders until the team's rate card + split matrix arrive; **wired into hours by Step 9** (deal fee → hrs by level, spread over the deal's dates or the group default duration) |
| **Actuals** | NetSuite time entries + matching (Step 4) | Paste area A Project code · B NetSuite employee · C week start · D hours (all input); calc E matched deal · F matched person · G level · H archetype · I planned hrs/wk · J weekly variance · K status. Name-mapping table at N/O (NetSuite name → roster person). Summary counts (matched/unmatched) at top. Ships with a flagged fabricated 20-row sample |
| **Variance** | Estimate vs actual (Step 4) | Per-deal (A code · B planned hrs/wk · C actual hrs · D planned over logged wks · E variance · F actual/planned) + calibration tables by level and by archetype (actual/planned ratio → estimate runs high <1 / low >1). Counts only fully-matched (status = OK) rows |
| **Settings** | All knobs | B3 capacity window start (Monday) · B4 probability-weighting toggle Yes/No · B8:B14 default hrs/wk by level (placeholders, fallback when no archetype) · **B19 Check-in as-of date (update weekly) · B20 stale-deal threshold in weeks (Step 5, placeholder = 4)** · columns D–L: dropdown source lists (named ranges) |
| **Review** | Data-health panel + migration items + status-mapping table | Top: **data-health panel** (Step 6) — unknown-person/unknown-code/duplicate-code integrity checks (OK/CHECK) + unstaffed-deals count. Then migration items: topic / person / project / detail / suggested action / **your decision** (yellow), in two sections ① needs a human, ② resolved from source |

**Core mechanics** (already implemented in the prototype script):

- Hours flow: `Assignments!I = override if set, else Settings default for the staffed-as level; 0 if Inactive` → `K = I × deal probability` (only when Settings!B4 = "Yes") → Capacity cell = SUMPRODUCT over assignments where the week falls inside the deal's date window (blank dates = all weeks) → Roster/Deals roll-ups via SUMIFS/COUNTIFS.
- Per-person involvement is **Assignments!G Active?**, distinct from deal lifecycle — this is how "pencils down for me, deal still live" is modeled. Deal lifecycle Dead/Closed/On hold/Internal forces assignments Inactive.
- Migration mapping: WIP→Active, Proposal→Proposal, Dead/Closed as-is, "General Code"→Internal, "Not active"→On hold; the 34 personal statuses map to Active/Inactive per the table on the Review tab; unknown strings → Active + review flag (conservative).
- Color code: blue text = input · black = formula · green = cross-sheet pull · yellow fill = fill/review this · grey row = inactive.
- One EXAMPLE row on Deals (`EXAMPLE_0`) and one Inactive example on Assignments; both deletable, counted nowhere.

## 5. Delivery roadmap

Model guidance in one line: **Opus** for steps where the design is still open or correctness is subtle (formula architecture, data-matching edge cases); **Sonnet** for well-specified execution against this plan and the existing script. Every session must follow the working agreements in §6 — in particular the hard stop after each step: deliver, report, and wait for the user's go-ahead before anything further.

| # | Step | Model | Size |
|---|------|-------|------|
| 1 | Finish & verify tracker v2 generation | **Sonnet** | Small |
| 2 | Apply team review round | **Sonnet** | Small |
| 3 | Time-phased effort templates | **Opus** | Large |
| 4 | NetSuite actuals import + variance | **Opus** | Medium |
| 5 | Weekly check-in tab | **Sonnet** | Small |
| 6 | Hardening & polish | **Sonnet** | Small |
| 7 | Scenario toggle (optional) | **Opus** | Medium |

### Step 1 — Finish & verify tracker v2 generation · **Sonnet**

`scripts/build_workbook.py` already exists and runs: it migrated **313 engagements** and **687 assignment rows**, applied **209** personal statuses, and produced **32 review items**. Whole-column references initially stalled LibreOffice recalculation; they are now bounded, but **recalc verification has not yet passed** — that is this step's first task.

- Regenerate the workbook, run the xlsx skill's `recalc.py` (timeout ≥ 300 s), fix until **zero formula errors**.
- Spot-check against independently computed values: pick 3 people, hand-count their Active assignments and committed hrs/wk from the Assignments tab and compare to Roster E/F; check one Capacity cell equals the sum of that person's weighted hours; check one deal's # staffed.
- Verify dropdowns (Person, Project ID, lifecycle, Active?) and conditional formatting work in a real Excel/LibreOffice open, and the Guide tab reads correctly.
- Acceptance: clean recalc JSON; spot-checks match; file opens with working validation; counts match the numbers above (± any deliberate fixes).
- Output: `workbook/US_MA_Tax_Scheduling_Tracker_v2_<date>.xlsx` committed, plus any script fixes.

### Step 2 — Apply team review round · **Sonnet** (after humans answer)

**Part A — data-recovery pass (done 2026-07-19, no human input needed).** Before the team touches anything, everything the *source data itself* can answer was resolved via the build script, shrinking the review list from 32 undifferentiated items to a 7-item human shortlist:

- Recovered all 22 "missing client" engagements' clients from the source Summary/Filtered tabs (data recovery, not invention — e.g., Brick_1 → Fengate, EV3 → Kalos LLP).
- Auto-resolved `EV3`/`EV3_1` (different clients — Kalos LLP vs Eva Equity Partners — a coincidental code collision, *not* a duplicate) and `Cosmetic _1` (both register rows same client, one engagement).
- Retriaged the Review tab into **①  Needs your decision (7)** and **②  Resolved from source (27, audit trail)**.
- Re-verified: 0 formula errors, capacity math unchanged. The workbook still holds no human input, so this was done by regenerating from the script (still legitimate — see §6 rule 2 nuance).

**Part B — apply the team's answers (done 2026-07-19 for the answers given).** The team filled in the 7-item Review shortlist. Applied: Patch MIF_1 → Active (the one capacity change), Patch CLS_1 → Inactive, weekly capacity 40 h/wk confirmed. Recorded/kept-as-is: no dates yet; proposal-probability semantics confirmed (50% default kept); the "40 h/wk" answer to the level-default question is the weekly-capacity figure, not a per-deal rate, so directional level-defaults were kept (and Step 3 archetypes supersede them anyway). **Still open** (noted on the Review tab): which-Kyle on 4 of the 5 projects (no Summary status existed, both Kyles kept per register), and the real per-level/per-archetype hour numbers.

- Applied by encoding the decisions in the generator (the human input was exactly 7 fully-reproducible cells — verified by diff), keeping the script authoritative through the Step 3 build-out; see §6 rule 2 for the cutover.
- Verified: 0 formula errors, Patch 9→10 active / 72→80 h/wk.

### Step 3 — Time-phased effort templates · **Opus** ✅ done 2026-07-19

Delivered — the flat hrs/wk model now has an estimation + phasing layer:

- **Templates tab**: 8 deal archetypes (Buy-side TDD, TDD + Structuring, Sell-side/VDD, Structuring-only, Tax equity/PW&A, FIRPTA/cross-border, Modeling/QoT, Ad-hoc) → hrs/wk per person **by level** + a 2-phase shape (front "diligence" fraction `f` at intensity `m1`, tail at `m2`, with `m2` a formula so the duration-weighted average is 1.0 — phasing redistributes hours, never adds/removes them). All numbers are placeholders (yellow) with a Review item.
- **Deals!Effort archetype** (dropdown). When set, the deal's per-level hrs/wk drive its assignments' planned hours (an Override still wins); the fallback when no archetype is the old Settings level-default.
- **Capacity is genuinely time-phased**: with an archetype **and** dates, a deal's weekly load follows the curve (front weeks heavier, tail lighter); missing either, it stays flat across the window (identical to before).
- **Design honored:** INDEX/MATCH-era only; **no 687×26 matrix** — phasing uses ~10 per-assignment helper columns + a 2-SUMIFS Capacity cell (§7); template numbers flagged.
- **Verified:** 0 errors across 27,002 cells; invariant confirmed (no archetype ⇒ numbers identical to Step 2); synthetic archetype+dates test on B2V_1 hand-matched — Senior-Associate rate 10→14, front week 19.6 (14×1.4), tail 8.4 (14×0.6), past-end 0. Guide updated.

### Step 4 — NetSuite actuals import + variance · **Opus** ✅ done 2026-07-19

Delivered — the actuals feedback loop:

- **CSV contract** documented on the Guide and the Actuals tab: `Project code | Employee (NetSuite) | Week start (Mon) | Hours`.
- **Actuals tab**: paste area (A–D) + per-row matching — deal (project code in Deals), person (roster direct-match, else the NetSuite→roster **name-mapping table** at N/O), level, archetype, the assignment's planned hrs/wk, and weekly variance. A live summary counts matched / unmatched-code / unmatched-person; unmatched rows are flagged red, never silently dropped.
- **Variance tab**: per-deal estimate vs actual, plus **calibration by level and by archetype** (actual ÷ planned over the logged weeks → >1 means the estimate runs low, <1 high), counting only fully-matched (status = OK) rows.
- **Verified against a flagged fabricated 20-row sample** (18 OK + 1 bad code + 1 unmapped person): 0 errors across 31,287 cells; matching, per-row/per-deal variance, and by-level calibration all hand-matched exactly (e.g. SA actual 62 / planned 60 → 1.03; VP 42/40 → 1.05; unmatched rows correctly excluded from calibration but counted in the report). Steps 1–3 numbers unchanged (Step 4 only adds tabs). **The sample must be cleared before real data is pasted** — and once actuals are pasted, that's the operational-data-at-scale cutover to bootstrap-only (§6 rule 2).

### Step 5 — Weekly check-in tab · **Sonnet** ✅ done 2026-07-19

Delivered — a formula-only agenda for the 15–20 min governance meeting, as a new **Check-in** tab:

- **§1 Capacity (next 4 weeks):** one row per roster person (row-aligned with Roster/Capacity, same pattern Capacity already uses to mirror Roster) — max load and max utilization over the first 4 Capacity week-columns, flagged Over-allocated (>100%) or Under-utilized (<50%), reusing the same thresholds already established on Roster/Capacity's conditional formatting.
- **§2 Deals needing attention:** one row per Active/Proposal deal (row-aligned with Deals) with a combined text Flag built from five independent checks: **Unstaffed** (#staffed = 0), **MissingLevel** (the deal's archetype expects hours from a level — Templates!B–H > 0 — but nobody at that level is actively staffed; deals with no archetype are exempt, since there's no "expected" to compare against), **DefaultProbability** (still Proposal at exactly 50%), **NoDates**, and **Stale** (weeks since `Added on` ≥ the new `Settings!B20` threshold).
- **Prerequisite added:** `Deals!Z` "Added on" (migrated rows = 2026-07-09 — the tracker's build date, not the deal's actual start, flagged as an assumption). New `Settings!B19` (check-in as-of date, blue/yellow — the team updates this before each meeting, since no volatile `TODAY()` is allowed) and `Settings!B20` (stale-weeks threshold, placeholder = 4).
- **Definitions judgment call:** "missing levels" isn't specified further in the brief; interpreted as "the archetype's expected level mix isn't fully staffed" (leveraging Step 3's Templates data rather than inventing a separate expected-mix model) and stated explicitly on the tab + a Review item, so the team can correct it if they meant something else.
- **Verified — the "manual filter" acceptance test, literally:** independently recomputed both agenda blocks from raw Assignments/Deals/Settings data (matching the workbook's own probability-weighting) and diffed against the sheet — **0 mismatches across all 23 people and all 306 Active/Proposal deals.** 0 formula errors across 35,269 cells. Since the live data has no archetypes set and nothing old enough to be stale, MissingLevel and Stale were *inert by default* (an untested-path risk per §7's Step-3 lesson) — proven on a synthetic scratch copy instead: setting one deal's archetype to expect a level nobody was staffed at correctly produced `MissingLevel`, and pushing the as-of date past the threshold correctly produced `Stale` on migrated deals. Guide's check-in how-to now says "open the Check-in tab."

### Step 6 — Hardening & polish · **Sonnet** ✅ done 2026-07-19

Delivered:

- **Sheet protection, formulas locked / inputs unlocked.** All 11 sheets protected (no password — a guardrail, not security; Review → Unprotect Sheet is always available). Every input cell (blue/yellow, including the blank spare rows in each input column) is unlocked; every formula/reference cell is locked, so a stray keystroke can't overwrite a formula. Insert/delete/**sort** rows are disabled on purpose — Capacity/Check-in/Variance mirror Deals & Roster row-for-row, so reordering would desync them; the team adds data in the pre-provided spare rows. Format/filter/select stay allowed. Verified per-sheet: sampled input cells read `locked=False`, formula cells `locked=True`, and the `formulas` engine still evaluates clean under protection.
- **Data-health panel** (top of the Review tab, recomputes live): assignments pointing at an unknown person or unknown project code, duplicate project codes, and Active/Proposal deals with nobody staffed. Integrity checks show green OK / red CHECK; unstaffed count is informational (points to Check-in). **All three integrity checks are green on the live file** (0/0/0); proven to actually fire by injecting a bogus person, a bogus code and a duplicate code on a scratch copy (→ 1/1/2, all CHECK). Caught a real bug in the process: `COUNTIF(range,"?")` treats `?` as a single-char wildcard, so the literal-`?` match needed escaping to `"~?"`.
- **Print setup on Check-in:** landscape, fit-to-width, header rows repeated, print area over the agenda (filter the Flag column first, then print).
- **Polish:** Guide gains a protection note in the color-code section and the new tabs in the how-to; column widths/wrapping pass.
- **Acceptance met:** every §4 edit path still works under protection (input cells editable, dropdowns intact); health integrity checks green on the live file; 0 errors across 35,295 cells; Steps 1–5 numbers unchanged.

### Step 7 — Scenario analysis · **Opus** ✅ done 2026-07-19 (on request)

Delivered — what-if analysis, deliberately **isolated** from the live plan (in Excel this is the feature most likely to confuse, so isolation was the design priority over a stateful toggle):

- **Scenario override inputs on Deals** (violet columns AA/AB/AC): scenario probability, scenario start, scenario end. Blank = use live. A scenario probability of **0%** (drop a deal) is a valid, distinct value from blank — so the effective-scenario values are computed on Deals with direct same-row `IF(blank, live, override)` refs (AD–AG), which the Assignments block pulls cleanly (probability preserves 0; dates use the 0→"" safe-lookup).
- **Assignments scenario mirror** (AB–AN): a full parallel copy of the live 2-segment phasing block, sourced from the effective-scenario values and reusing the archetype m1/m2. So the scenario re-phases correctly if scenario dates are set, not just re-weights.
- **Scenario tab** (read-only): a per-person impact summary (Live peak vs Scenario peak over the next 4 weeks, Change, "NEWLY OVER" flag) + a full scenario heatmap + team Live-vs-Scenario weekly totals. No stateful Settings toggle — the overrides *are* the levers, and with none set the tab equals Live (every Change = 0), which is far less confusing than a mode switch that silently changes what the live-looking views show.
- **Isolation verified two ways:** (1) with no overrides, the scenario heatmap equals Capacity cell-for-cell, all 23 impact Changes = 0, team delta = 0, and every Steps-1–6 live number is unchanged; (2) setting B2V_1's scenario probability to 0% dropped exactly Lauren −5 and Jennifer Wu −10 (their planned hrs on it) in the Scenario tab while **Live stayed identical** (Deals planned still 15, their live capacity unchanged) and an unrelated person (Yiyi) didn't move. 0 errors across 49,337 cells.
- **On §10 (web app trigger):** the scenario works, but it took ~17 parallel helper columns (Assignments AB–AN + Deals AA–AG) and is the workbook's most complex machinery — a fair signal that *further* analytical depth (multi-scenario compare, saved scenarios, Monte-Carlo over probabilities) is where a real app would pay off. Flagged for the team, not acted on.

**This completes the original Excel roadmap (Steps 1–7).** Feedback Round 1 (§5A) extends it with Steps 8–11.

## 5A. Feedback Round 1 — power-user feedback (Justine, July 21) → Steps 8–11

**Source:** Justine Morin's email of July 21, 2026 — the tracker's primary user; a working session on this feedback is scheduled for the same evening. The original is in French; the five points, faithfully rendered:

1. **Too many tabs.** Cut the workbook to the essentials. (It currently has 12.)
2. **Workstream mapping per engagement**, with her exact taxonomy (wording below is verbatim from the email and must be preserved in the dropdowns):
   - **Tax Due Diligence — deal type** (pick one): Asset Deal (or DRE) · Partnership with push out · Partnership no push out · C Corporation · S Corporation
   - **Target footprint:** Domestic only (Y/N)
   - **Tax Due Diligence Scope** (pick one): Inquiry basis · Limited Procedures · Key Findings
   - **Other TDD relevant scoping items:** Carve out · Non-US jurisdiction coordination · 8A / Taxand · N/A
   - **Modeling** (each can be ticked or N/A — multi-select): Corporate tax modeling (sub-choice: Basic (no roll up) · Building roll up) · Tracking Model Review · Property tax Modeling · SALT Modeling · Tax Credit Assumption Review · Step Up Calculation · Section 382 limitation · FIRPTA Modeling · Other modeling complexities · N/A
   - **Structuring:** Strawman deck · Structure paper · Opinion
   - **Legal docs review**
3. **Fee quote → allocation.** Enter each engagement's fee quote and have the allocation flow from it. She explicitly wants to brainstorm the mechanism ("on peut brainstorm sur ce point ce soir").
4. **Tentative timelines with weekly-adjusting flags.** Timelines are the hardest thing to pin down — set tentative kick-off and delivery dates and have the flag adjust each week.
5. **Shareable visual output.** The output should be easy to understand and share with the D/MD group; the back-end "matrix" tab she is happy to update separately.

**How this lands on the existing design** — three lucky breaks and one open problem:

- The legacy attribute columns (`Deals!K–P`: transaction type, entity class, scope, industry, complexity, timeline) are **empty for every real deal** (verified July 21 — only the EXAMPLE row has values; they were always yellow to-fill placeholders). Her taxonomy can therefore *replace* K–P outright with zero data loss. Re-verify emptiness in-session before replacing (rule 5).
- Point 4 is largely the existing Check-in machinery: every flag already recomputes when the weekly as-of date (`Settings!B19`) is bumped — one cell, once a week, deterministic (no volatile `TODAY()`, per §7). The work is richer flag semantics and surfacing, not a new engine.
- Points 1 and 5 solve each other: one **Dashboard** tab as the shareable front door, back-end tabs **hidden, never deleted** (they keep computing; unhide any time).
- Point 3's mechanism was the one open problem — **resolved at the July 21 review** (see Decisions below): fees per workstream drive the allocation, via a rate card and a level-split matrix the team supplies.

**Decisions — recorded July 21 (user's answers to the 5-item checklist):**

1. **Essential tabs — DECIDED:** visible = **Dashboard · Deals · Assignments** as proposed; everything else hidden but live (including Roster — unhide anytime).
2. **Fee → allocation — DECIDED (a fourth mechanism, superseding options A/B/C):** fees are entered **per workstream** on each engagement and **drive the hour allocation**. The team supplies two parameter sets: (a) the **rate card by level** ($/hr) and (b) an **illustrative % split by level per workstream** (how each workstream's fee divides across levels). Both are *awaited inputs* — yellow placeholders until provided.
3. **Timeline flags — DECIDED:** flag set and thresholds exactly as proposed (`Tentative (no dates) · Not started · Kick-off ≤1 wk · In flight · Delivery ≤2 wks · Overdue · Delivered`); weekly deterministic bump of `Settings!B19`, no volatile `TODAY()`.
4. **Estimation driver — DECIDED:** the workstreams **replace** the 8 effort archetypes; archetype machinery kept hidden as legacy (rule 5), no longer the driver.
5. **Live copy — DECIDED:** no manual edits exist in the live copy; the repo snapshot is authoritative and Step 8 may regenerate safely. (Cutover watch continues: the first real data entry — fees, workstream ticks, dates — flips the generator to bootstrap-only per §6 rule 2.)

**Computed fee→hours design** (the mechanics implied by decision 2 — confirm the four ★ interpretation points at Step 8 kickoff):

```
hours(level, workstream)   = fee(workstream) × split%(workstream, level) ÷ rate(level)
deal hours by level        = Σ over that deal's priced workstreams
deal hrs/wk by level       = deal hours by level ÷ duration in weeks
per-assignment rate        = deal hrs/wk at their staffed-as level ÷ # active assignees at that level
```

- ★ **Fee granularity:** fees quoted per **workstream group** — TDD · Modeling · Structuring · Legal docs review (· Other) — one fee column each on Deals, total fee as a formula. (Per sub-item — e.g. each of the 9 modeling ticks — would be ~18 fee columns; assumed too granular for quoting.) The scoping ticks say *what's in scope*; the group fee says *how big*.
- ★ **Duration:** hrs/wk uses the kick-off→delivery span (Step 10's dates); when dates are blank, a **default duration** knob in Settings (placeholder) so fee-driven hours still land somewhere visible rather than silently vanishing.
- ★ **Person allocation:** a level's hrs/wk splits **evenly** across the active assignees at that level on the deal; an assignment Override still wins (same `Assignments!I` choke-point contract as today); a level with expected hours but nobody staffed surfaces on Check-in/Dashboard as MissingLevel **with its unallocated hrs/wk shown**.
- ★ **Probability weighting** stays where it is (`Assignments!K`), applied after all of the above — unchanged choke point.

**Awaited inputs before Step 9 can compute real numbers:** the rate card by level · the split-% matrix (workstream group × level, each row summing to 100%) · then per-deal fees as the team enters them. Steps 8–11 are otherwise unblocked; **dev work starts only on the user's explicit "go"**.

### Step 8 — Workstream scoping model · **Opus** · Large ✅ done 2026-07-19

Point 2, delivered. (Decision 5 confirmed no live-copy edits, so the generator regenerated safely — no diff needed.)

- **Retired `Deals!K–P` in place** (relabelled "(retired)", dropdowns removed, columns hidden — not deleted, rule 5; re-verified empty for every real deal, and grep-confirmed no formula referenced them) and **appended the scoping taxonomy at AH–AX** with Justine's wording verbatim: TDD deal type · Domestic only? (Y/N) · TDD scope · Other TDD scoping (incl. N/A) · nine Modeling columns (Corporate tax modeling as a `Basic (no roll up) / Building roll up / blank=N/A` dropdown, the other eight `Y`/blank) · three Structuring `Y`/blank · Legal docs review `Y`/blank. New dropdown lists on Settings M–Q. **Appending (not shifting) means zero churn to the ~30 downstream Deals-column references — the safe choice.**
- **Reworked Templates into the fee-allocation parameter tab:** rate card $/hr by level, the split-% matrix (5 workstream groups × 7 levels, each row `=100%` checksum with red CF on drift), and default durations per group — all yellow placeholders + a Review item. The archetype table was kept exactly in place (rows 2–10, still referenced by the phasing formulas) and relabelled legacy.
- **Verified:** taxonomy headers match the email exactly; K–P hidden; split checksums all evaluate to 100%; **the invariant holds — with no fees/workstreams entered, every live number is unchanged** (Patch 80, Yiyi 150, B2V_1 15, health 0/0/0), because Step 8 only added data columns + parameter tables and touched no existing formula. 0 errors across 49,485 cells. The hours math (fees → hours) lands in Step 9.

### Step 9 — Fees → hour allocation · **Opus** · Medium/Large ✅ done 2026-07-21

Point 3, per the recorded decision + the computed design above. Built against the **placeholder** rate card + split matrix — the engine is correct the instant the team's real numbers replace them (no rework):

- `Deals` gained **fee inputs per workstream group** (AY–BC: TDD · Modeling · Structuring · Legal docs review · Other), a **fee-total gate** (BD), per-group **duration helpers** (BE–BI = kick-off→delivery weeks, ≥1, or the group default when dates are blank), and **fee hrs/wk by level** (BJ–BP = Σ fee × split% ÷ rate ÷ duration). BE–BP are hidden calc helpers.
- Engine wired through the **`Assignments!I` choke point**: Override → **fee-driven** → archetype → level default. Fee-driven per assignment = `AP` = level hrs/wk ÷ active assignees at that level; `AO` = the deal's total fee (the has-fee gate). K/weighting untouched. The seg1 window now also honours a fee-priced deal's dates so its hours spread **flat across kick-off→delivery** (the Scenario mirror got the identical treatment, so Live≡Scenario still holds with no overrides).
- **Deferred to Step 11:** a *visible* "unstaffed funded level" flag (Dashboard). The behaviour is already correct — those hours compute on `Deals` but land on nobody — and the caveat is documented in Guide + Review.
- **Verified:** 0 errors across 57,771 cells; **invariant holds** — with no fees entered, every live number is byte-identical to Step 8 (Justine 160, Yiyi 150, Patch 80; B2V_1 15; team 2093.5), because the engine is gated on a fee > 0. Two synthetic scratch tests hand-checked end-to-end against independent recomputation: (1) fees → level hours (`fee×split÷rate÷dur`) → per-person rates, date-windowing (in-window − out-of-window = the exact fee contribution), and the gate leaving un-priced deals at their level defaults; (2) even split across two same-level people (30→15+15), an unstaffed funded level (hours compute, land on nobody), and per-level conservation (Σ per-person = deal-weekly = fee×split÷rate÷dur).

### Step 10 — Timeline flags · **Sonnet** · Small ✅ done 2026-07-21

Point 4, delivered per the recorded decision (flag set + wording verbatim from decision 3):

- **Renamed `Deals!I/J` headers** to **Kick-off date / Delivery date** — the cells, their formats, and every formula referencing them are untouched (the phasing, fee-duration, and scenario logic all keep working; zero churn, verified).
- **Added `Deals!BQ` "Timeline status"** (visible calc column) — 7 mutually-exclusive branches off `Settings!B19` (the existing weekly as-of date; no volatile `TODAY()`): `Tentative` (a date is blank) → `Not started` (>7 days to kick-off) → `Kick-off ≤1 wk` (≤7 days) → `In flight` → `Delivery ≤2 wks` (≤14 days to delivery) → `Overdue` (past delivery, lifecycle still Active) → `Delivered` (past delivery, any other lifecycle). **Surfaced on Check-in §2** as a new column (Overdue rows highlight red); the Step 11 Dashboard will pull the same column.
- **Verified:** 0 errors; **invariant holds** — every real deal has a blank kick-off or delivery date, so all show `Tentative` and no capacity/fee number moves (baseline spot checks unchanged). A synthetic date-matrix scratch test drove **all 7 branches** (independent Python transcription of the spec, not the formula), and **bumping the as-of date +7 days** moved the flags exactly as recomputed. Boundary semantics documented in-code: `≤1 wk`/`≤2 wks` are literal 7-/14-day thresholds; Overdue vs Delivered split on lifecycle = Active.

### Step 11 — Dashboard + tab diet · **Opus** · Medium

Points 1 + 5. A new **Dashboard** tab in position 1 — the D/MD share view: headline tiles (team utilization next 4 weeks, # over-allocated, # deals needing attention, upcoming deliveries), a compact capacity heat strip, and the timeline-flag list; print/PDF-ready, fully protected, zero inputs. Then the tab diet **per the recorded decision**: visible = **Dashboard · Deals · Assignments**; the other ten (Guide, Capacity, Check-in, Scenario, Roster, Templates, Actuals, Variance, Settings, Review) hidden via `sheet_state="hidden"` — never deleted, still computing; Guide gains an "unhiding tabs" note; Check-in's content is absorbed into the Dashboard. Acceptance: the file opens on Dashboard; hidden tabs still drive every number (0 errors, invariants hold); a PDF of the Dashboard is legible standalone; §4 gains a visibility column.

**Order:** 8 → 9 → 10 → 11 (fees lean on the workstream mix; the Dashboard shows everything). Step 10 is independent and can slot in anywhere. All §6 working agreements still apply — one step per session, verify with the §7 engine + independent hand-checks, hard stop after each delivery.

## 6. Working agreements for every Claude session

1. **Hand the session:** this `PLAN.md`, `scripts/build_workbook.py`, the **latest** workbook from `workbook/`, and (for Steps 1–2) the original `Scheduling_US_MA_July_9.xlsx`.
2. **Regeneration cutover (refined in Step 2/3).** The generator stays authoritative *only while every cell is reproducible from the source + a small set of encoded decisions*. That held through Step 3: the team's Step 2 input was exactly 7 cells (diff-verified), encoded in `REVIEW_DECISIONS`/`DECISION_TEXT`, so regenerating lost nothing and made Step 3's structural build safe to verify. **The cutover to bootstrap-only comes the moment per-deal operational data is entered at scale** — dates, probabilities, archetypes, overrides typed straight into the workbook — because that is impractical to back-encode. After that, edit in place (load without `data_only`, preserve formulas), save a dated copy first, and never regenerate. When unsure whether that threshold has passed, assume it has.
3. **Always verify:** run the xlsx skill's `recalc.py` (timeout ≥ 300 s) until zero errors — **or, if it hangs in your sandbox (it does here), the `formulas` pure-Python engine** (§7). Then hand spot-check 2–3 computed values against an independent recomputation, and exercise any date/phase feature on a synthetic-dated scratch copy (§7). A clean run proves formulas evaluate, not that they're right.
4. **Never invent business numbers silently.** Any placeholder (hours, capacity, probability) gets yellow fill + a Review-tab item.
5. **Snapshot per step:** commit `workbook/US_MA_Tax_Scheduling_Tracker_v2_<date>.xlsx` (or the step's output) and the updated scripts to the repo. The live team copy sits on SharePoint/Teams; the repo holds the spec, scripts and snapshots.
6. Keep the §4 tab/column contract stable; if a step must change it, update §4 in the same commit.
7. **One step per session; hard stop after delivery.** Complete the step's acceptance criteria, commit and push, then stop with a handoff report (delivered / verified / open questions / next step). Never roll into the next step unprompted — the pause is where the team reviews the workbook and answers the open questions the next step depends on.

## 7. Technical notes for implementers (hard-won, read before writing formulas)

- **Bounded ranges only.** Whole-column references (`Assignments!$A:$A`) are far too slow to evaluate at this workbook's size; bounded (`$A$2:$A$769`) is the standard. Current extents: Deals rows 2–355, Assignments rows 2–769, Roster rows 2–40 — spare rows are pre-filled with guarded formulas (`IF($A2="","",…)`).
- **Function whitelist:** Excel-2007-era only — SUMIFS/COUNTIFS/SUMPRODUCT/INDEX/MATCH/IFERROR. No XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE (they break the LibreOffice verification harness and older Excel). If TEXTJOIN/IFS/SWITCH/MAXIFS/MINIFS are ever needed, write them as `_xlfn.TEXTJOIN(…)` etc.
- **Capacity uses SUMIFS against sentinel-date helper columns, not SUMPRODUCT.** `Assignments!N/O` ("Eff. start/end") mirror `L/M` (Deal start/end) but substitute `DATE(1900,1,1)` / `DATE(2100,12,31)` when blank, so `Capacity!<cell> = SUMIFS(Assignments!$K$2:$K$769, Assignments!$A$2:$A$769, person, Assignments!$G$2:$G$769, "Active", Assignments!$N$2:$N$769, "<="&week, Assignments!$O$2:$O$769, ">="&week)`. An earlier SUMPRODUCT version (5 multiplied boolean arrays × 768 rows × 1,014 cells) was functionally fine but too slow to verify at all in some environments (see next bullet) — SUMIFS against pre-resolved sentinel columns is both correct and fast, and is the pattern to extend if Step 3 adds more time-phased math.
- **`INDEX`/`MATCH` into a genuinely blank cell returns the number `0`, not `""`.** This bit us directly: `Assignments!L` (deal start) does `INDEX(Deals!$I..., MATCH(...))`, and since every migrated deal has a blank `Deals!Expected start`, every `L` cell silently evaluated to `0` — not blank. The `N = IF(L="", DATE(1900,1,1), L)` sentinel guard never fired (0 ≠ ""), so `N` became `0` too, which fails `>=week` for every real week and **zeroed out the entire Capacity view**. It was caught only by independently hand-computing expected values and diffing against the sheet — a clean recalc (no `#REF!`/`#VALUE!`) would never have flagged it, since 0 is a valid number. Fixed via `safe_lookup()` in `build_workbook.py`, which treats a found-but-zero result the same as not-found. **Any new `INDEX`/`MATCH` pulling from a column that can legitimately be blank (dates, free-text fields) must use this pattern** — wrap the lookup so a `0` result collapses to `""`, don't just guard the pre-lookup blank check.
- **Time-phasing (Step 3) is two SUMIFS over pre-multiplied segment columns — NOT a 687×26 matrix.** The curve is 2-piece: a front phase `[eff-start, split]` at intensity `m1` and a tail `(split, eff-end]` at `m2`, where `split = start + f·(end−start)` and `m2` is derived so `f·m1+(1−f)·m2 = 1` (hours conserved). Each assignment carries helper columns: `S/T` = the deal's `m1/m2` (pulled from Deals, default 1), `U` = split date, and two segment triples `V/W/X` = seg1 start/end/hrs and `Y/Z/AA` = seg2 start/end/hrs, where `X = K·m1`, `AA = K·m2`. **The whole thing keys off one flag: `U` (split date) is non-blank iff the deal has both dates AND an archetype.** When `U=""`: seg1 = `[1900,2100]` at `hrs=K` (flat, all weeks) and seg2 is an empty range (`start 2100 > end 1900`), so a non-phased row behaves exactly as the pre-Step-3 flat model. Capacity cell = `SUMIFS(seg1_hrs, person, "Active", V"<="wk, W">="wk) + SUMIFS(seg2_hrs, …, Y"<="wk, Z">="wk)`. This adds ~10 columns to Assignments and doubles the SUMIFS count — cheap, and it kept the workbook at 27,002 cells / a clean sub-minute `formulas`-engine verify. The 2D archetype-rate lookup (`Assignments!I`) is `INDEX(Templates!$B$3:$H$N, MATCH(archetype, names), MATCH(staffed-as, level-headers))` — evaluated twice (a `>0` guard then the value); keep the Templates level headers (`B2:H2`) exactly equal to the roster level strings or the `MATCH` returns `#N/A`.
- **Verify Step-3-style features on a synthetic-dated scratch copy.** Every migrated deal has blank dates+archetype, so phasing is inert in the live file and a "no errors" run proves nothing about it. Assign one archetype + start/end on a throwaway copy and hand-check the assignment's seg1/seg2 hrs (`K·m1`, `K·m2`) and a front-vs-tail capacity cell — that is the only thing that actually exercises the curve.
- **Data validation uses named ranges** (`RosterNames`, `DealCodes`, `LifecycleList`, `ArchetypeList`, …) — portable across Excel versions and LibreOffice, unlike direct cross-sheet DV references.
- **Heatmap zero-hiding** via number format `0.0;-0.0;` keeps the grid readable.
- **`COUNTIF`/`SUMIF`/`COUNTIFS` treat `?` and `*` in the criterion as wildcards** (`?` = any single char, `*` = any run). The health panel counts assignments whose level/client lookup returned the literal string `"?"` (unknown ref) — written bare, `COUNTIF(range,"?")` matches every single-character cell and silently returns a huge count. Escape the literal with a tilde: `"~?"`. (Caught only because the health panel was value-checked against a known-clean file — it read 768 instead of 0.) Any COUNTIF/SUMIF whose criterion is a literal containing `?`/`*` needs the `~` escape.
- **Protection (Step 6):** all sheets carry `SheetProtection(sheet=True)`, no password. Inputs are unlocked by explicit 1-based `(col,row,col,row)` ranges (see `unlock_specs` in `build_workbook.py`) that include the blank spare rows; everything else stays locked. Insert/delete/**sort** rows are disabled (openpyxl semantics: attribute `True` = action *disabled*) because Capacity/Check-in/Variance mirror Deals & Roster row-for-row. When adding a column of formulas, remember new cells default to `locked=True` — add their input columns to `unlock_specs` if users must type there. Protection does not affect the `formulas`/LibreOffice engines.
- The probability toggle flows through exactly one choke point: `Assignments!K`. Change weighting logic there only.
- Dates are hardcoded inputs (window start = Settings!B3), never `TODAY()`/`NOW()` — volatile functions would make recalc results shift between sessions.

### Verification: recalc.py may not work in your sandbox — have a fallback ready

The xlsx skill's `recalc.py` (LibreOffice headless) **hung indefinitely in this session's container**, reproducibly, even on a trivial one-formula file — confirmed to be an environment issue (near-zero CPU consumed across 5-minute hangs; a fresh/cold LibreOffice profile is the trigger, confirmed via direct `soffice` invocation outside Python's `subprocess`, and even `--convert-to` — a code path with no macro involved at all — failed the same way). This is very likely container-specific, not universal — **try `recalc.py` first, with a real timeout (≥300s)**, before assuming it's broken for you too.

If it hangs or errors out, the fallback that worked cleanly here: `pip install formulas` (pure-Python Excel engine, no LibreOffice dependency). It loaded and calculated this entire workbook (16,613 cells) in under a minute:

```python
import formulas
xl = formulas.ExcelModel().loads("workbook/<file>.xlsx").finish()
sol = xl.calculate()
# keys look like "'[<FILENAME-EXACT-CASE>.xlsx]<SHEETNAME-UPPER>'!<CELL-UPPER>"
# sol[key].value is usually a 1x1 array — unwrap with .ravel().tolist()[0]
```

Scan `sol` for the standard error tokens (`#VALUE!`, `#REF!`, `#NAME?`, `#DIV/0!`, `#NULL!`, `#NUM!`, `#N/A`) the same way `recalc.py` would. **This does not bake cached values into the delivered `.xlsx`** (unlike `recalc.py`, which rewrites the file in place) — the file still opens with formulas uncalculated until Excel/LibreOffice computes them on open (standard, automatic behavior for any real user opening it; only automated tools reading `data_only=True` without opening the file first would see blanks). If baked-in cached values turn out to matter for a later step, that's an open problem — `formulas` doesn't write back, and `recalc.py` is what actually rewrites the file.

Whichever engine passes, **a clean run only proves formulas evaluate without error — it does not prove they're right** (the blank-lookup bug above produced zero errors and silently wrong numbers). Always pair it with independent hand-computation on 3–4 sampled cells from the raw input data, not from the formulas themselves — and if a feature has never been exercised by the actual data (e.g., date-phased capacity, when every migrated deal has blank dates), test it separately on a throwaway scratch copy with synthetic values rather than trusting that "no errors" means "this code path works."

## 8. Governance (how the team runs it)

Weekly 15–20 min (VP/Director/MD/Partner): open Capacity → who's red/amber next 4 weeks; Check-in tab (from Step 5) → unstaffed incoming deals, proposals to re-probability, dates to fill; record changes live in the workbook. Monthly: compare NetSuite actuals vs estimates (Step 4 views) and tune the level defaults / templates.

## 9. Open data questions (feed Step 2)

Associates — real or planned? · Renewable pool: blended into M&A staffing or separate? · Real weekly capacity + billable targets by level · PTO source · Probability buckets · Do register Project IDs match NetSuite codes exactly? · NetSuite history depth · Confidentiality walls on any deals? · Non-deal time (BD, proposals, internal) as pseudo-engagements?

## 10. Deferred: web app

Revisit when the workbook hits its natural limits — signals: concurrent-edit conflicts on SharePoint, the team outgrowing ~35–40 people, real appetite for scenario analysis (Step 7 demand), or the NetSuite CSV ritual becoming a burden. The workbook's tab/column contract (§4) is deliberately database-shaped so migration stays a straight import.
