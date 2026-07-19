# CLAUDE.md

## What this project is

An Excel-based scheduling and capacity tracker for Leo Berwick's US M&A Tax team (~23 people, ~290 live engagements): deals → staffing assignments → hours-by-level estimates → a person×week capacity heatmap, with a NetSuite actuals feedback loop. **Everything is delivered as an Excel workbook — there is no application code.** The web app idea is deferred (PLAN.md §10).

**Read `PLAN.md` before doing anything.** §4 is the workbook spec (tab/column contract), §5 the step roadmap, §6 working agreements, §7 technical notes. Follow the step protocol below — one step per session, hard stop after delivery.

**Current status:** Step 2 — *data-resolvable portion done* (`workbook/US_MA_Tax_Scheduling_Tracker_v2_2026-07-19.xlsx`, supersedes the same-day Step 1 file). Recovered all 22 missing clients from the source Summary/Filtered tabs; auto-resolved EV3/EV3_1 (different clients, not a duplicate) and Cosmetic_1 (one engagement); retriaged the Review tab into **7 items that need a human** + 27 resolved-from-source (audit trail). Re-verified: 0 formula errors across 16,694 cells, capacity numbers unchanged (client names don't touch the math). **The workbook still contains NO human edits, so the build script is still authoritative and regeneration is still safe.** Step 2 *proper* — applying the team's 7 decisions + real capacities / level-defaults / probabilities / dates — is **still waiting on those inputs from Justine's group**; the moment they are entered, the script becomes bootstrap-only (rule 2).
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
2. **Never regenerate the workbook once it contains human input the script can't reproduce.** Regeneration is safe *only* while every cell is script-derivable from the source file (true through the data-recovery part of Step 2). The moment anyone enters a non-derivable decision — which-Kyle, real capacities, tuned level-defaults, per-proposal probabilities, deal dates, a Review-tab answer — the build script becomes **bootstrap-only**: from then on, edit the workbook in place (load *without* `data_only=True`, preserve formulas) and save a dated copy before editing. When in doubt about whether human input has landed, assume it has and edit in place.
3. **Verify before delivering:** run the xlsx skill's `recalc.py` with timeout ≥ 300 s until **zero formula errors**. If it hangs/fails in your sandbox (it did in this one — see PLAN.md §7), fall back to `pip install formulas` (pure-Python, no LibreOffice) the same way that section documents. Either way, a clean run only proves formulas *evaluate* — always also hand spot-check 3–4 computed values against an independent recomputation from the raw input data (not from the formulas themselves), and if a feature has never been exercised by the actual data (e.g., date-phased capacity when every current deal has blank dates), test it separately on a throwaway scratch copy with synthetic values. This is not optional box-checking: Step 1 shipped a bug (§7) that produced zero formula errors and silently zeroed the entire Capacity view, caught only by this kind of independent check.
4. **Never invent business numbers silently.** Placeholders (hours, capacities, probabilities, template values) get blue text + yellow fill + a Review-tab item.
5. **Nothing is deleted in migrations or edits** — rows get flagged (Inactive, grey, Review item), not removed. Preserve the raw source text columns.
6. Don't touch `Assignments!K` weighting logic except deliberately — it is the single choke point for probability weighting.

## Excel engineering constraints

- **Functions:** Excel-2007-era only — SUMIFS, COUNTIFS, SUMPRODUCT, INDEX, MATCH, IFERROR. **Never** XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE (they break the LibreOffice verification harness and older clients). TEXTJOIN/IFS/SWITCH/MAXIFS/MINIFS only with the `_xlfn.` prefix.
- **Bounded ranges only.** Whole-column refs (`$A:$A`) are too slow at this size. Current extents: Deals 2–355, Assignments 2–769, Roster 2–40.
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
- Hours unit is **hours/week per person per deal**; defaults by level live at `Settings!B8:B14` and are placeholders until the team tunes them.

## Git & delivery conventions

- Branch: `claude/advisory-resource-scheduling-ccg6t3`. Push with `git push -u origin <branch>`.
- One commit per step minimum: the step's workbook snapshot + any script changes + updated docs (PLAN.md §4 if the contract moved, the Current-status line here).
- Workbook files are binary — always save a new dated snapshot rather than relying on git to diff them.
- This repo contains confidential client data; it must stay private. Don't paste client/deal names into PRs, issues, or anything that leaves the repo.
