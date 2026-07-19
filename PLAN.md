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
| **Deals** | One row per engagement | A Project ID · B client as filed · C primary client · D end client · E referral? · F lifecycle · G phase · H probability · I/J expected start/end · K–P attributes (txn type, entity class, scope, industry, complexity, timeline) · Q # staffed (formula) · R planned hrs/wk (formula) · S source status · T notes · **U effort archetype (input, dropdown)** · V/W front/tail intensity (f) · X front fraction (f) · Y phase split date (f) |
| **Assignments** | One row per person × deal; the staffing ledger | A person · B roster level (formula) · C project ID · D client (f) · E deal lifecycle (f) · F staffed-as level · G **Active?** · H override hrs/wk · I planned hrs/wk used (f: override → **archetype rate** → level default) · J deal probability (f) · K weighted hrs/wk (f) · L/M deal start/end (f) · N/O effective start/end (f, sentinel dates when L/M blank) · P source status · Q notes · **R–AA phasing helpers** (archetype, front/tail x, split date, seg1 start/end/hrs, seg2 start/end/hrs — see §7) |
| **Roster** | Team list | A person · B level · C specialty · D weekly capacity hrs · E active assignments (f) · F committed hrs/wk (f) · G utilization (f, CF flags) |
| **Templates** | Deal archetypes → effort (Step 3) | A archetype name · B–H hrs/wk per person by level (Partner→Associate + Renewable) · I typical duration (wks) · J front-phase fraction · K front intensity ×· L tail intensity × (formula: keeps duration-weighted avg = 1). All B–K placeholders (yellow) |
| **Settings** | All knobs | B3 capacity window start (Monday) · B4 probability-weighting toggle Yes/No · B8:B14 default hrs/wk by level (placeholders, fallback when no archetype) · columns D–L: dropdown source lists (named ranges) |
| **Review** | Migration items needing sign-off + the status-mapping table | Topic / person / project / detail / suggested action / **your decision** (yellow). Two sections: ① needs a human, ② resolved from source |

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

### Step 4 — NetSuite actuals import + variance · **Opus**

- Define the CSV contract for a NetSuite saved search: `project code, person, week start, hours` (document it on the Guide tab for whoever builds the search).
- An **Actuals** tab with a paste-in area; matching by project code + a person-name mapping table (NetSuite display names ≠ tracker first names — reuse the roster as the mapping anchor).
- An unmatched-rows report (codes or people that don't match get listed, not silently dropped).
- Estimate-vs-actual views: per deal (planned vs burned by level) and a calibration summary per level/archetype suggesting default adjustments.
- Acceptance: with a fabricated 20-row sample CSV (clearly marked fake), matching, variance and unmatched reporting all verifiably correct; recalc clean.

### Step 5 — Weekly check-in tab · **Sonnet**

A formula-only agenda for the 15–20 min governance meeting:

- Over-allocated people (next 4 weeks, from Capacity), under-utilized people, Active/Proposal deals with **# staffed = 0** or missing levels, proposals still at the default 50%, deals with blank dates, stale rows.
- Prerequisite: append an `Added on` date column to Deals (migrated rows = 2026-07-09) so aging is computable.
- Acceptance: each agenda block cross-checked against a manual filter; recalc clean; Guide's check-in how-to updated to "open the Check-in tab".

### Step 6 — Hardening & polish · **Sonnet**

- Sheet protection with input cells unlocked (formula columns can't be typed over accidentally); strict data validation where safe (reject unknown names/codes).
- A small health-check panel: assignments pointing at unknown people/codes, deals with no staffing, duplicate codes.
- Print/export area for the check-in; final pass on widths, wrapping, Guide wording.
- Acceptance: protected sheets still allow every intended edit path in §4's color code; health checks all green on the live file; recalc clean.

### Step 7 — Scenario toggle (optional) · **Opus**

Only if the team asks for it after using the tracker: scenario override columns on Deals (probability / dates) plus a Settings switch so Capacity can show Live vs Scenario side by side. In Excel this is the feature most likely to add confusing complexity — build it only on demand, and consider whether its arrival is really the trigger to revive the web app (§10).

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
