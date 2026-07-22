# CLAUDE.md

## What this project is

An Excel-based scheduling and capacity tracker for Leo Berwick's US M&A Tax team (~23 people, ~290 live engagements): deals → staffing assignments → hours-by-level estimates → a person×week capacity heatmap, with a NetSuite actuals feedback loop. **Everything is delivered as an Excel workbook — there is no application code.** The web app idea is deferred (PLAN.md §10).

**Read `PLAN.md` before doing anything.** §4 is the workbook spec (tab/column contract), §5 the step roadmap, §6 working agreements, §7 technical notes. Follow the step protocol below — one step per session, hard stop after delivery.

**Current status: Steps 1–11 delivered & verified — Feedback Round 1 COMPLETE** (`workbook/US_MA_Tax_Scheduling_Tracker_v2_2026-07-19.xlsx`, 13 tabs, 3 visible; 0 errors across 59,969 cells at last full verify). **Step 11 done — Dashboard + tab diet:** new **Dashboard** tab (first + active, zero inputs, fully protected) — 5 KPI tiles (team utilization 4wk · # over-allocated · # deals needing attention · # deliveries ≤2 wks · # overdue), a next-8-weeks per-person capacity heat strip (mirrors Capacity), and a compact deals-needing-attention list rank-extracted from Check-in without FILTER (running-rank helper on Check-in cols K/L + INDEX/MATCH on the Dashboard, "+N more" overflow note). **Tab diet:** visible = Dashboard·Deals·Assignments; the other ten hidden via `sheet_state="hidden"` (never deleted, still computing; Guide + Dashboard header explain unhide). **Deferred Step 9 flag landed:** `FeeLevelGap` token (a fee funds a level nobody's staffed at) on Check-in §2 + Dashboard, silent until fees entered. **Invariant verified:** hidden tabs still drive every number (baseline spot checks unchanged); the Dashboard's own cells independently re-derived (util 227% by design with blank dates, over-alloc 19, attention 17; heat strip mirrors Capacity; attention list matches Check-in ranked order). **All of Justine's 5 points delivered.** Awaited team inputs before numbers are "real": rate card + split-% matrix (Templates), per-deal fees, Kick-off/Delivery dates. No further steps queued. Older open items: which-Kyle on 4 projects; NetSuite go-live; "missing level"/"stale" definitions.
*(Update this line in the same commit whenever a step completes.)*

## Repo map

- `PLAN.md` — spec + roadmap. The §4 tab/column contract is stable; changing it requires updating §4 in the same commit.
- `scripts/build_workbook.py` — regenerates the tracker from the July 9 source file. **Bootstrap only: forbidden after Step 2** (it would destroy the team's live edits).
- `workbook/` — dated output snapshots (`US_MA_Tax_Scheduling_Tracker_v2_<YYYY-MM-DD>.xlsx`). Created in Step 1. After Step 2, the latest snapshot (or the team's SharePoint copy the user provides) is the source of truth — edit it in place with openpyxl; never rebuild.
- `data/` — (optional) the source inputs: `Scheduling_US_MA_July_9.xlsx` and the June 4 requirements email. The repo owner decides whether to commit these (confidential client data); if absent, ask the user to upload them for Steps 1–2. Later steps only need the latest workbook.

## Step protocol — one step per session, hard stop after delivery

1. Before touching anything, confirm which PLAN.md §5 step this session is executing.
2. Deliver that step: meet its acceptance criteria, commit + push (workbook snapshot, script changes, doc updates), and update the Current-status line above in the same commit.
3. **Then STOP.** End with a handoff report: what was delivered, verification evidence (recalc result + the spot-checks you ran), any decisions or questions for the humans (e.g., open Review-tab items), and which step comes next. Do **not** begin the next step — not partially, not "while we're here", not even if the request was ambiguous about scope — until the user explicitly asks in a new instruction. The pause between steps is where the team reviews the workbook and answers open questions; work done past it is likely to be thrown away.

## Non-negotiable rules

1. **Load the `xlsx` skill before touching any workbook.** `openpyxl`/`pandas` may need `pip install` in a fresh container.
2. **Regeneration cutover.** The generator stays authoritative *only while every cell is reproducible from the source + a small set of encoded decisions*. This held through Step 3: Step 2's human input was exactly 7 cells (diff-verified), encoded in `REVIEW_DECISIONS`/`DECISION_TEXT`, so regenerating lost nothing. **The cutover to bootstrap-only is when per-deal operational data gets entered at scale** — dates, probabilities, archetypes, overrides typed straight into the workbook — which can't be back-encoded. After that: edit in place (load *without* `data_only=True`, preserve formulas), save a dated copy first, never regenerate. When in doubt whether that threshold has passed, assume it has and edit in place.
3. **Verify before delivering:** run the xlsx skill's `recalc.py` with timeout ≥ 300 s until **zero formula errors**. If it hangs/fails in your sandbox (it did in this one — see PLAN.md §7), fall back to `pip install formulas` (pure-Python, no LibreOffice) the same way that section documents. Either way, a clean run only proves formulas *evaluate* — always also hand spot-check 3–4 computed values against an independent recomputation from the raw input data (not from the formulas themselves), and if a feature has never been exercised by the actual data (e.g., date-phased capacity when every current deal has blank dates), test it separately on a throwaway scratch copy with synthetic values. This is not optional box-checking: Step 1 shipped a bug (§7) that produced zero formula errors and silently zeroed the entire Capacity view, caught only by this kind of independent check.
4. **Never invent business numbers silently.** Placeholders (hours, capacities, probabilities, template values) get blue text + yellow fill + a Review-tab item.
5. **Nothing is deleted in migrations or edits** — rows get flagged (Inactive, grey, Review item), not removed. Preserve the raw source text columns.
6. Don't touch `Assignments!K` weighting logic except deliberately — it is the single choke point for probability weighting.

## Excel engineering constraints

- **Functions:** Excel-2007-era only — SUMIFS, COUNTIFS, SUMPRODUCT, INDEX, MATCH, IFERROR. **Never** XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE (they break the LibreOffice verification harness and older clients). TEXTJOIN/IFS/SWITCH/MAXIFS/MINIFS only with the `_xlfn.` prefix.
- **Bounded ranges only.** Whole-column refs (`$A:$A`) are too slow at this size. Current extents: Deals 2–355, Assignments 2–769, Roster 2–40, Templates 3–10.
- **Time-phasing (Step 3) = 2 SUMIFS over pre-multiplied segment columns, never a 687×26 matrix.** A 2-piece curve (front `[start,split]`×`m1`, tail `(split,end]`×`m2`, `m2` derived so hours conserve) lives in Assignments helper cols `S–AA`; it keys off `U` (split date), which is non-blank iff the deal has both dates AND an archetype, and collapses to the old flat behaviour otherwise. See PLAN.md §7 before touching it.
- **Guarded spare rows are load-bearing.** Every row inside a formula extent carries `IF($A2="","",…)` formulas — don't clear those cells or shrink ranges past them.
- **Capacity is SUMIFS against sentinel-date helper columns** (`Assignments!N/O`, "Eff. start/end" — mirror `L/M` but substitute 1900-01-01/2100-12-31 when blank), not SUMPRODUCT. See PLAN.md §7 if extending this.
- **`INDEX`/`MATCH` into a blank cell returns `0`, not `""`.** Any lookup pulling from a column that can legitimately be blank (dates, free-text) must collapse a found-but-zero result to `""` (see `safe_lookup()` in `build_workbook.py`) — a plain `IF(result="",...)` guard on the caller side will not catch it, and this class of bug produces zero formula errors while being silently wrong. This exact bug shipped once in Step 1 (zeroed the whole Capacity view) before independent verification caught it.
- **No volatile functions** (`TODAY()`, `NOW()`) — results must not shift between sessions. The capacity window start is the input at `Settings!B3`.
- **Dropdowns via named ranges** (`RosterNames`, `DealCodes`, `LifecycleList`, `PhaseList`, `TxnTypeList`, `EntityClassList`, `ScopeList`, `TimelineList`, `ActiveList`, `YesNoList`, `LevelList`) — portable, unlike direct cross-sheet DV refs.
- **Formatting:** Arial 10 throughout. Color code — blue text = input · black = formula · green = cross-sheet pull · yellow fill = fill/review this · grey row = inactive. Heatmap hides zeros via number format `0.0;-0.0;`. Dates `yyyy-mm-dd`, probabilities `0%` stored as fractions.

## Domain quick reference

- **Levels:** Partner → MD → Director → VP → Senior Associate (→ Associate, none on roster yet), plus a cross-cutting 4-person **Renewable Specialist** pool (Dorian, Yiyi, Eric, Robbie) doing high-volume/low-hours review work.
- **Name collisions (why free text was banned):** Kyle Kidd = Partner, Kyle Risser = Director; Will B. W. = Partner, Will Covalt = Director. In the old file the level column a name sat in was the only disambiguator; in the tracker, people are canonical Roster entries.
- **Deal lifecycle vs personal involvement are different fields:** `Deals!Lifecycle` (Proposal/Active/On hold/Pens down/Complete–to invoice/Invoiced/Closed/Dead/Internal) says where the deal is; `Assignments!Active?` says whether that person still spends time. "Pencils down for me, deal still live" = lifecycle Active + assignment Inactive.
- **Probability prefills:** WIP → 100%, Proposal → 50% (yellow, to review). Blank deal dates → the deal counts in **every** capacity week until dates are filled (by design, to force the data in).
- Hours unit is **hours/week per person per deal**. Source of an assignment's rate: an **Override** (`Assignments!H`) wins; else the deal's **effort archetype** (`Deals!U` → `Templates`) rate for that level; else the flat `Settings!B8:B14` level default. All archetype/default numbers are placeholders until the team tunes them.
- **Effort archetypes** (`Templates` tab, Step 3): 8 deal types → hrs/wk by level + a diligence-heavy→lighter phase shape. A deal with an archetype **and** dates gets time-phased in Capacity; missing either, it's flat. Phasing conserves total hours (it redistributes across the deal's weeks).

## Git & delivery conventions

- Branch: `claude/advisory-resource-scheduling-ccg6t3`. Push with `git push -u origin <branch>`.
- One commit per step minimum: the step's workbook snapshot + any script changes + updated docs (PLAN.md §4 if the contract moved, the Current-status line here).
- Workbook files are binary — always save a new dated snapshot rather than relying on git to diff them.
- This repo contains confidential client data; it must stay private. Don't paste client/deal names into PRs, issues, or anything that leaves the repo.
