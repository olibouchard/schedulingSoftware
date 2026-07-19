#!/usr/bin/env python3
"""Build the US M&A Tax scheduling tracker (workbook v2) from the July 9 source file.

Reads the hand-maintained staffing workbook, cleans it (canonical names, controlled
statuses, merged duplicates), and emits a structured tracker with:
  Guide - how to use, color legend, assumptions
  Capacity - person x week planned-hours heatmap with over-allocation flags
  Deals - one row per engagement (attributes, probability, dates, dropdowns)
  Assignments - one row per person x deal (hours driven by level defaults, overridable)
  Roster - canonical team list with capacity and utilization
  Settings - dropdown lists, level defaults, capacity window
  Review - migration decisions that need human sign-off

Usage: python3 build_workbook.py <source.xlsx> <output.xlsx>
"""

import re
import sys
import datetime as dt
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- constants

LEVELS = ["Partner", "MD", "Director", "VP", "Senior Associate",
          "Renewable Specialist", "Associate"]

# Canonical roster, from the side columns of the July 9 file (level -> people).
ROSTER = OrderedDict([
    ("Partner", ["Tania", "John", "Will B. W.", "Kyle Kidd"]),
    ("MD", ["Zack", "Ethan", "Jeremy", "Jason"]),
    ("Director", ["Justine", "Kyle Risser", "Will Covalt", "Lauren"]),
    ("VP", ["George", "Spencer", "Patch"]),
    ("Senior Associate", ["Vidhi", "Nina", "Colleen", "Jennifer Wu"]),
    ("Renewable Specialist", ["Dorian", "Yiyi", "Eric", "Robbie"]),
])

# Sheet1 staffing columns -> level
COL_LEVEL = {5: "Partner", 6: "MD", 7: "Director", 8: "VP",
             9: "Senior Associate", 10: "Renewable Specialist"}

# First-name resolution inside a level column (handles the two Kyles / Wills)
LEVEL_ALIASES = {
    "Partner": {"kyle": "Kyle Kidd", "will": "Will B. W.", "william": "Will B. W.",
                "tania": "Tania", "john": "John"},
    "Director": {"kyle": "Kyle Risser", "will": "Will Covalt", "william": "Will Covalt",
                 "justine": "Justine", "lauren": "Lauren"},
}
GLOBAL_ALIASES = {"william b.": "Will B. W.", "will b. w.": "Will B. W.",
                  "jennifer": "Jennifer Wu"}

LIFECYCLES = ["Proposal", "Active", "On hold", "Pens down",
              "Complete - to invoice", "Invoiced", "Closed", "Dead", "Internal"]
LIFE_RANK = {v: i for i, v in enumerate(
    ["Active", "Proposal", "On hold", "Pens down", "Complete - to invoice",
     "Invoiced", "Closed", "Dead", "Internal"])}
PHASES = ["Pre-LOI", "Diligence", "Signing to close", "Post-close"]
TXN_TYPES = ["Stock", "Asset", "Partnership", "Mixed", "N/A"]
ENTITY_CLASSES = ["Corp", "Flow-through", "Mixed", "N/A"]
SCOPES = ["TDD", "Structuring", "Modeling", "TDD + Structuring", "Full scope",
          "PW&A / Review", "SALT", "Other"]
TIMELINES = ["Standard", "Compressed"]

# PLACEHOLDER defaults - hours/week one person typically spends on one deal, by
# level. To be tuned by the team (Settings tab, yellow cells).
DEFAULT_HOURS = OrderedDict([
    ("Partner", 2), ("MD", 3), ("Director", 5), ("VP", 8),
    ("Senior Associate", 10), ("Renewable Specialist", 2), ("Associate", 12),
])
DEFAULT_CAPACITY = 40          # PLACEHOLDER weekly capacity hours per person
PROB_ACTIVE, PROB_PROPOSAL = 1.0, 0.5   # prefill; proposals to be reviewed
N_WEEKS = 26
SPARE_DEALS, SPARE_ASSIGN, ROSTER_LAST = 40, 80, 40

# Per-person engagement status (Summary tab) -> assignment Active/Inactive
STATUS_ACTIVE = {
    "ongoing", "ongoing / jeremy", "ongoing ad-hod", "ongoing--close to invoicing",
    "tdd done / structuring ongoing", "closing", "finished phase i", "salt deal",
    "waiting greenlight to kick off diligence.", "proposal", "proposal sent",
    "proposal - to be coming soon",
}
STATUS_INACTIVE = {
    "closed", "dead", "died", "invoiced", "invoicing soon", "pencils down",
    "penicls down", "pens down", "pens dowm", "on pause",
    "not involved in ongoing work", "not really involved in ongoing work",
    "ongoing but not involved with", "closed to be invoiced",
    "closed--wire for payment should have been already sent",
    "client stopped responding to us re: invoicing.",
    "this was probably when i was helping check tax insurance datarooms.",
}
STATUS_REVIEW_INACTIVE = {
    "not sure what's going on with this one. this is an old project.",
    "don't know what this is",
}

# ---- Human review decisions (from the team's filled-in Review tab, 2026-07-19).
# Encoded here so the generator reproduces the human-decided state exactly; the
# workbook stays script-reproducible through the Step 3 build-out. Only decisions
# that change a cell value live here - confirmations of existing state don't.
REVIEW_DECISIONS = {
    # (person, project) -> Active? flag, an explicit per-person involvement call
    ("Patch", "MIF_1"): ("Active", "Active - confirmed by review 2026-07-19 (user: yes)"),
    ("Patch", "CLS_1"): ("Inactive", "Inactive - confirmed by review 2026-07-19 (user: no)"),
}
# What the team wrote in the yellow decision column, echoed back on the Review tab
# next to each of the 7 shortlist items so the tab records the resolution.
DECISION_TEXT = {
    "kyle": "Kyle Kidd on CopiaAdHoc_1 (per review). Note: the Summary tab held "
            "no involvement status for any of these 5, so nothing needed applying "
            "- both Kyles stay Active per the register. Which Kyle on Dorsia_1 / "
            "Gateway_3 / Guardian_3 / Shika_1: still open.",
    "dates": "No start/end dates yet (per review). Deals keep counting in every "
             "week until dates are added; the Step 3 phase curve activates then.",
    "leveldefault": "Kept directional placeholders. The user's '40 h/wk' answer is "
                    "the weekly CAPACITY (applied below); a 40 h/wk-per-deal default "
                    "would over-allocate massively. Step 3 archetype templates now "
                    "drive per-level hours - tune those instead.",
    "prob": "Semantics confirmed: probability = chance the proposal converts to WIP. "
            "Kept the 50% default pending per-proposal estimates.",
    "capacity": "Confirmed 40 h/wk per person (real, no longer a placeholder).",
    "cls": "Patch INACTIVE on CLS_1 (user: no) - applied.",
    "mif": "Patch ACTIVE on MIF_1 (user: yes) - applied.",
}

# ---------------------------------------------------------------- styling

ARIAL = "Arial"
F_TITLE = Font(name=ARIAL, size=14, bold=True, color="1F3864")
F_HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=ARIAL, size=10)
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")        # blue = input
F_LINK = Font(name=ARIAL, size=10, color="008000")         # green = cross-sheet
F_NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_YELLOW = PatternFill("solid", fgColor="FFFF00")       # user should fill/review
FILL_GREY = PatternFill("solid", fgColor="F2F2F2")
FILL_RED = PatternFill("solid", fgColor="F8CBAD")
FILL_AMBER = PatternFill("solid", fgColor="FFE699")
FILL_BLUE = PatternFill("solid", fgColor="D9E1F2")
THIN_BTM = Border(bottom=Side(style="thin", color="D9D9D9"))

WROTE = []


def norm(s):
    return re.sub(r"\s+", " ", str(s).strip()) if s is not None else ""


def person_key(raw, level):
    """Resolve a raw staffing-cell name to a canonical roster person."""
    n = norm(raw).lower()
    if not n:
        return None
    if n in LEVEL_ALIASES.get(level, {}):
        return LEVEL_ALIASES[level][n]
    if n in GLOBAL_ALIASES:
        return GLOBAL_ALIASES[n]
    for lvl, people in ROSTER.items():
        for p in people:
            if p.lower() == n:
                return p
    return None


# ---------------------------------------------------------------- extraction


def extract(src):
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["Sheet1"]
    deals = OrderedDict()          # code -> dict
    assigns = OrderedDict()        # (person, code) -> dict
    review = []                    # (topic, person, project, detail, action)
    unmatched_names = []

    person_level = {p: lvl for lvl, ps in ROSTER.items() for p in ps}

    for r in range(4, ws.max_row + 1):
        code = norm(ws.cell(row=r, column=3).value)
        if not code:
            continue
        client = norm(ws.cell(row=r, column=2).value)
        raw_status = norm(ws.cell(row=r, column=4).value)
        life = {"WIP": "Active", "Proposal": "Proposal", "Dead": "Dead",
                "Closed": "Closed", "General Code": "Internal",
                "Not active": "On hold"}.get(raw_status, "Active")
        if raw_status not in ("WIP", "Proposal", "Dead", "Closed"):
            review.append(("Unusual register status", "", code,
                           f"'{raw_status}' mapped to lifecycle '{life}'",
                           "Confirm or correct the Lifecycle on the Deals tab"))
        if code in deals:                                  # duplicate register row
            prev = deals[code]
            keep = "Active" if "Active" in (prev["life"], life) else life
            same_client = client.lower() == prev["client"].lower()
            prev["source"] += f" + {raw_status} (duplicate rows merged)"
            prev["life"] = keep
            detail = ("Two rows in the July 9 register"
                      + (", same client" if same_client else ", DIFFERENT clients")
                      + f"; differ mainly in status - kept lifecycle '{keep}'")
            action = ("Confirmed one engagement; no action needed"
                      if same_client else
                      "Different clients - check whether these are two deals")
            review.append(("Duplicate register rows merged", "", code, detail,
                           action))
        else:
            parts = [norm(p) for p in client.split(" : ") if norm(p)]
            referral = "Yes" if "(referral source)" in client.lower() else "No"
            primary = re.sub(r"\(referral source\)", "", parts[0],
                             flags=re.I).strip() if parts else ""
            deals[code] = dict(code=code, client=client, primary=primary,
                               end=parts[-1] if len(parts) > 1 else "",
                               referral=referral, life=life, source=raw_status)
        for col, lvl in COL_LEVEL.items():
            cell = ws.cell(row=r, column=col).value
            if not cell:
                continue
            for part in re.split(r"[,/&]| and ", str(cell)):
                if not norm(part):
                    continue
                p = person_key(part, lvl)
                if p is None:
                    unmatched_names.append((norm(part), lvl, code))
                    continue
                key = (p, code)
                if key not in assigns:
                    assigns[key] = dict(person=p, code=code, staffed_as=lvl,
                                        active="Active", raw="", note="")
                if person_level.get(p) != lvl:
                    assigns[key]["note"] = (f"Listed under {lvl} but roster level "
                                            f"is {person_level.get(p)}")

    # ---- per-person statuses from the Summary tab
    ws2 = wb["Summary by Person"]
    summary_person_map = {**{p: p for ps in ROSTER.values() for p in ps},
                          "William B.": "Will B. W."}

    # Recover client names for engagements that never reached the register, from
    # the Summary / Filtered tabs (both carry a client column). This is data
    # recovery from the source, not invention.
    proj_client_recovery = {}
    for tab in ("Summary by Person", "Filtered Status"):
        if tab not in wb.sheetnames:
            continue
        wsx = wb[tab]
        for rr in range(2, wsx.max_row + 1):
            pj = norm(wsx.cell(row=rr, column=3).value)
            cl = norm(wsx.cell(row=rr, column=2).value)
            if pj and cl and cl not in proj_client_recovery.get(pj, []):
                proj_client_recovery.setdefault(pj, []).append(cl)

    cur, applied, kyle_unresolved = None, 0, []
    for r in range(2, ws2.max_row + 1):
        a, proj = ws2.cell(row=r, column=1).value, norm(ws2.cell(row=r, column=3).value)
        if a is not None:
            s = norm(a)
            try:
                float(s)
            except ValueError:
                if s:
                    cur = s
        if not proj or cur is None:
            continue
        wipprop = norm(ws2.cell(row=r, column=4).value)
        raw = norm(ws2.cell(row=r, column=5).value)
        person = summary_person_map.get(cur)
        if cur == "Kyle":                       # ambiguous block: resolve per deal
            hits = [p for p in ("Kyle Kidd", "Kyle Risser") if (p, proj) in assigns]
            person = hits[0] if len(hits) == 1 else None
            if person is None:
                kyle_unresolved.append(proj)
                continue
        if person is None:
            continue
        low = raw.lower()
        flag, needs_review = None, False
        if low in STATUS_ACTIVE:
            flag = "Active"
        elif low in STATUS_INACTIVE:
            flag = "Inactive"
        elif low in STATUS_REVIEW_INACTIVE:
            flag, needs_review = "Inactive", True
        elif raw:
            flag, needs_review = "Active", True
        if proj not in deals:                   # work that never reached the register
            life = ("Closed" if low in ("closed",) else
                    "Proposal" if wipprop == "Proposal" or "proposal" in low
                    else "Active")
            rec_clients = proj_client_recovery.get(proj, [])
            recovered = rec_clients[0] if rec_clients else ""
            multi = f" (source also lists: {'; '.join(rec_clients[1:])})" if len(rec_clients) > 1 else ""
            deals[proj] = dict(code=proj, client=recovered, primary=recovered,
                               end="", referral="No", life=life,
                               source=f"Summary tab only ({wipprop or 'n/a'})")
            if recovered:
                review.append((
                    "Engagement recovered from Summary tab", person, proj,
                    f"Not in the register; found on the Summary tab (status: "
                    f"'{raw or wipprop}'). Client recovered from source: "
                    f"{recovered}{multi}.",
                    "Client filled in - confirm, then add deal attributes "
                    "(type / scope / complexity) when known"))
            else:
                review.append((
                    "Engagement missing from register", person, proj,
                    f"Found only on the Summary tab (status: '{raw or wipprop}'); "
                    f"no client in source",
                    "Confirm it is real; add client and attributes on Deals tab"))
        if (person, proj) not in assigns:
            assigns[(person, proj)] = dict(person=person, code=proj,
                                           staffed_as=person_level.get(person, ""),
                                           active="Active", raw="",
                                           note="Added from Summary tab")
        rec = assigns[(person, proj)]
        rec["raw"] = raw
        if flag:
            rec["active"] = flag
            applied += 1
        if needs_review:
            rec["note"] = (rec["note"] + "; " if rec["note"] else "") + "Status unclear"
            review.append(("Unclear personal status", person, proj, f"'{raw}'",
                           "Set Active? on the Assignments tab"))

    if kyle_unresolved:
        review.append(("Ambiguous 'Kyle' rows skipped", "Kyle (?)",
                       ", ".join(sorted(set(kyle_unresolved))[:12]),
                       "Summary tab has one 'Kyle' block; these projects have "
                       "neither/both Kyles on the register so the status was not applied",
                       "Check which Kyle and set Active? manually"))
    for nm, lvl, code in unmatched_names:
        review.append(("Name not in roster", nm, code,
                       f"Appears in the {lvl} column but matches nobody on the roster",
                       "Add the person to Roster or fix the name, then add an "
                       "Assignments row"))
    if "EV3" in deals and "EV3_1" in deals:
        c1, c2 = deals["EV3"]["primary"], deals["EV3_1"]["primary"]
        if c1 and c2 and c1.lower() != c2.lower():
            review.append((
                "Name collision - resolved", "", "EV3 / EV3_1",
                f"Different clients (EV3 = {c1}; EV3_1 = {c2}) - coincidental "
                f"code similarity, not a duplicate",
                "Kept as two separate engagements; no action needed"))
        else:
            review.append(("Possible duplicate engagement", "", "EV3 / EV3_1",
                           "Justine tracks 'EV3', Nina tracks 'EV3_1'",
                           "Merge into one code if they are the same deal"))

    # deals with a dead/closed lifecycle force their assignments inactive
    for (p, c), rec in assigns.items():
        if deals[c]["life"] in ("Dead", "Closed", "On hold", "Internal") \
                and rec["active"] == "Active":
            rec["active"] = "Inactive"
            rec["note"] = (rec["note"] + "; " if rec["note"] else "") + \
                f"Deal lifecycle is {deals[c]['life']}"

    # human review decisions win last (authoritative over any auto-derivation)
    for (person, proj), (flag, note) in REVIEW_DECISIONS.items():
        if (person, proj) in assigns:
            assigns[(person, proj)]["active"] = flag
            assigns[(person, proj)]["note"] = note
    return deals, assigns, review, applied


def safe_lookup(rr, last_d, src_range, notfound='""'):
    """INDEX/MATCH into Deals by Assignments!$C{rr} (project ID) that treats
    both 'not found' and 'found but the source cell is blank' as blank.
    Excel/Calc's INDEX returns the number 0 for a truly empty referenced
    cell, not "" - left unguarded, blank Deals dates/client cells would
    read back as 0 (a fake epoch date, or a literal "0" client name)."""
    idx = f'INDEX({src_range},MATCH($C{rr},Deals!$A$2:$A${last_d},0))'
    return (f'=IF($C{rr}="","",IFERROR(IF({idx}=0,"",{idx}),{notfound}))')


# ---------------------------------------------------------------- helpers


def style_header(ws, row, cols, widths=None):
    for i, title in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=title)
        c.font, c.fill = F_HDR, FILL_HDR
        c.alignment = Alignment(vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 28


def wcell(ws, row, col, value, font=F_BODY, fill=None, fmt=None, align=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    WROTE.append(c)
    return c


# ---------------------------------------------------------------- build


def build(deals, assigns, review, out_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    guide = wb.create_sheet("Guide")
    cap = wb.create_sheet("Capacity")
    wd = wb.create_sheet("Deals")
    wa = wb.create_sheet("Assignments")
    wr = wb.create_sheet("Roster")
    st = wb.create_sheet("Settings")
    rv = wb.create_sheet("Review")
    for sheet, color in [(guide, "808080"), (cap, "2E7D32"), (wd, "1F4E79"),
                         (wa, "1F4E79"), (wr, "1F4E79"), (st, "BF8F00"),
                         (rv, "C00000")]:
        sheet.sheet_properties.tabColor = color

    deal_rows = sorted(deals.values(), key=lambda d: (LIFE_RANK[d["life"]],
                                                      d["code"].lower()))
    roster_flat = [(p, lvl) for lvl, ps in ROSTER.items() for p in ps]
    roster_pos = {p: i for i, (p, _) in enumerate(roster_flat)}
    assign_rows = sorted(assigns.values(),
                         key=lambda a: (roster_pos.get(a["person"], 99),
                                        a["code"].lower()))

    n_deals = len(deal_rows) + 1                       # +1 example row
    n_assign = len(assign_rows) + 1
    LAST_D = 1 + n_deals + SPARE_DEALS                 # last formula row, Deals
    LAST_A = 1 + n_assign + SPARE_ASSIGN               # last formula row, Assignments

    # ---------------- Settings
    wcell(st, 1, 1, "Settings & Lists", F_TITLE)
    wcell(st, 3, 1, "Capacity window start (a Monday)", F_BOLD)
    wcell(st, 3, 2, dt.date(2026, 7, 20), F_INPUT, FILL_YELLOW, "yyyy-mm-dd")
    wcell(st, 4, 1, "Probability-weight the Capacity view? (Yes/No)", F_BOLD)
    wcell(st, 4, 2, "Yes", F_INPUT, FILL_YELLOW)
    wcell(st, 6, 1, "Default planned hours/week (one person, one deal), by level",
          F_BOLD)
    wcell(st, 6, 4, "PLACEHOLDER values set by Claude from the July 9 staffing "
                    "shape - not firm data. Tune them here; every assignment "
                    "without an override uses these.", F_NOTE)
    wcell(st, 7, 1, "Level", F_HDR, FILL_HDR)
    wcell(st, 7, 2, "Default hrs/wk", F_HDR, FILL_HDR)
    for i, (lvl, hrs) in enumerate(DEFAULT_HOURS.items(), start=8):
        wcell(st, i, 1, lvl)
        wcell(st, i, 2, hrs, F_INPUT, FILL_YELLOW, "0.0")
    lists = [("D", "Lifecycle", LIFECYCLES), ("E", "Phase", PHASES),
             ("F", "Transaction type", TXN_TYPES),
             ("G", "Entity classification", ENTITY_CLASSES),
             ("H", "Scope", SCOPES), ("I", "Timeline", TIMELINES),
             ("J", "Active flag", ["Active", "Inactive"]),
             ("K", "Yes/No", ["Yes", "No"]), ("L", "Levels", LEVELS)]
    for col, title, values in lists:
        ci = openpyxl.utils.column_index_from_string(col)
        wcell(st, 7, ci, title, F_HDR, FILL_HDR)
        for j, v in enumerate(values, start=8):
            wcell(st, j, ci, v)
    st.column_dimensions["A"].width = 44
    st.column_dimensions["B"].width = 14
    for col, _, _ in lists:
        st.column_dimensions[col].width = 20
    wcell(st, 17, 1, "Lists to the right feed the dropdowns. Add values as the "
                     "practice evolves; deleting a value in use does not change "
                     "existing rows.", F_NOTE)

    names = {
        "RosterNames": f"Roster!$A$2:$A${ROSTER_LAST}",
        "DealCodes": f"Deals!$A$2:$A${LAST_D}",
        "LifecycleList": f"Settings!$D$8:$D${7+len(LIFECYCLES)}",
        "PhaseList": f"Settings!$E$8:$E${7+len(PHASES)}",
        "TxnTypeList": f"Settings!$F$8:$F${7+len(TXN_TYPES)}",
        "EntityClassList": f"Settings!$G$8:$G${7+len(ENTITY_CLASSES)}",
        "ScopeList": f"Settings!$H$8:$H${7+len(SCOPES)}",
        "TimelineList": f"Settings!$I$8:$I${7+len(TIMELINES)}",
        "ActiveList": "Settings!$J$8:$J$9",
        "YesNoList": "Settings!$K$8:$K$9",
        "LevelList": f"Settings!$L$8:$L${7+len(LEVELS)}",
    }
    for nm, ref in names.items():
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

    # ---------------- Roster
    style_header(wr, 1, ["Person", "Level", "Specialty", "Weekly capacity (hrs)",
                         "Active assignments", "Committed hrs/wk", "Utilization"],
                 [16, 18, 14, 12, 11, 11, 10])
    r = 2
    for person, lvl in roster_flat:
        wcell(wr, r, 1, person, F_INPUT)
        wcell(wr, r, 2, lvl, F_INPUT)
        wcell(wr, r, 3, "Renewable" if lvl == "Renewable Specialist" else "Core M&A",
              F_INPUT)
        wcell(wr, r, 4, DEFAULT_CAPACITY, F_INPUT, FILL_YELLOW, "0.0")
        r += 1
    for rr in range(2, ROSTER_LAST + 1):
        wcell(wr, rr, 5, f'=IF($A{rr}="","",COUNTIFS(Assignments!$A$2:$A${LAST_A},'
                         f'$A{rr},Assignments!$G$2:$G${LAST_A},"Active"))',
              F_BODY, fmt="0")
        wcell(wr, rr, 6, f'=IF($A{rr}="","",SUMIFS(Assignments!$I$2:$I${LAST_A},'
                         f'Assignments!$A$2:$A${LAST_A},$A{rr},'
                         f'Assignments!$G$2:$G${LAST_A},"Active"))',
              F_BODY, fmt="0.0")
        wcell(wr, rr, 7, f'=IF(OR($A{rr}="",$D{rr}=""),"",IF($D{rr}=0,"",'
                         f'$F{rr}/$D{rr}))', F_BODY, fmt="0%")
        for ccol in range(1, 8):
            wr.cell(row=rr, column=ccol).border = THIN_BTM
    wr.freeze_panes = "A2"
    wr.conditional_formatting.add(
        f"G2:G{ROSTER_LAST}",
        FormulaRule(formula=["AND(ISNUMBER($G2),$G2>1)"], fill=FILL_RED,
                    stopIfTrue=True))
    wr.conditional_formatting.add(
        f"G2:G{ROSTER_LAST}",
        FormulaRule(formula=["AND(ISNUMBER($G2),$G2>0.85)"], fill=FILL_AMBER,
                    stopIfTrue=True))
    wr.conditional_formatting.add(
        f"G2:G{ROSTER_LAST}",
        FormulaRule(formula=["AND(ISNUMBER($G2),$G2<0.5)"], fill=FILL_BLUE,
                    stopIfTrue=True))
    dv_lvl = DataValidation(type="list", formula1="LevelList", allow_blank=True)
    wr.add_data_validation(dv_lvl)
    dv_lvl.add(f"B2:B{ROSTER_LAST}")
    wcell(wr, ROSTER_LAST + 2, 1,
          "Capacity 40h is a placeholder - set real weekly capacity (blue/yellow "
          "cells are inputs). Add new team members in the blank rows; they appear "
          "everywhere automatically.", F_NOTE)

    # ---------------- Deals
    hdr = ["Project ID", "Client (as filed)", "Primary client",
           "End client / target", "Referral?", "Lifecycle", "Phase",
           "Probability", "Expected start", "Expected end", "Transaction type",
           "Entity classification", "Scope", "Industry", "Complexity flags",
           "Timeline", "# staffed", "Planned hrs/wk", "Source status (Jul 9)",
           "Notes"]
    style_header(wd, 1, hdr, [22, 38, 26, 22, 8, 14, 13, 10, 11, 11, 13, 13, 14,
                              14, 16, 11, 8, 10, 16, 30])
    example = dict(code="EXAMPLE_0", client="Example Client LLC : Example Target",
                   primary="Example Client LLC", end="Example Target",
                   referral="No", life="Active", source="(example)")
    r = 2
    for d in [example] + deal_rows:
        is_ex = d["code"] == "EXAMPLE_0"
        live = d["life"] in ("Active", "Proposal")
        wcell(wd, r, 1, d["code"], F_INPUT)
        wcell(wd, r, 2, d["client"])
        wcell(wd, r, 3, d["primary"], F_INPUT)
        wcell(wd, r, 4, d["end"], F_INPUT)
        wcell(wd, r, 5, d["referral"], F_INPUT)
        wcell(wd, r, 6, d["life"], F_INPUT)
        wcell(wd, r, 7, "Diligence" if is_ex else None, F_INPUT,
              FILL_YELLOW if live and not is_ex else None)
        prob = (PROB_ACTIVE if d["life"] == "Active"
                else PROB_PROPOSAL if d["life"] == "Proposal" else 0.0)
        wcell(wd, r, 8, prob, F_INPUT,
              FILL_YELLOW if d["life"] == "Proposal" else None, "0%")
        wcell(wd, r, 9, dt.date(2026, 7, 20) if is_ex else None, F_INPUT,
              FILL_YELLOW if live and not is_ex else None, "yyyy-mm-dd")
        wcell(wd, r, 10, dt.date(2026, 10, 30) if is_ex else None, F_INPUT,
              FILL_YELLOW if live and not is_ex else None, "yyyy-mm-dd")
        for col, val in [(11, "Stock"), (12, "Corp"), (13, "TDD + Structuring"),
                         (14, "Renewables"), (15, "Cross-border"),
                         (16, "Standard")]:
            wcell(wd, r, col, val if is_ex else None, F_INPUT,
                  FILL_YELLOW if live and not is_ex else None)
        wcell(wd, r, 19, d["source"])
        if is_ex:
            wcell(wd, r, 20, "EXAMPLE row - shows the expected formats; delete "
                             "any time. Nobody is staffed on it.", F_NOTE)
        r += 1
    for rr in range(2, LAST_D + 1):
        wcell(wd, rr, 17, f'=IF($A{rr}="","",COUNTIFS(Assignments!$C$2:$C${LAST_A},'
                          f'$A{rr},Assignments!$G$2:$G${LAST_A},"Active"))',
              F_BODY, fmt="0")
        wcell(wd, rr, 18, f'=IF($A{rr}="","",SUMIFS(Assignments!$I$2:$I${LAST_A},'
                          f'Assignments!$C$2:$C${LAST_A},$A{rr},'
                          f'Assignments!$G$2:$G${LAST_A},"Active"))',
              F_BODY, fmt="0.0")
        for ccol in range(1, 21):
            wd.cell(row=rr, column=ccol).border = THIN_BTM
    wd.freeze_panes = "B2"
    wd.conditional_formatting.add(
        f"A2:T{LAST_D}",
        FormulaRule(formula=['OR($F2="Dead",$F2="Closed",$F2="Invoiced")'],
                    fill=FILL_GREY))
    for col, name in [("F", "LifecycleList"), ("G", "PhaseList"),
                      ("K", "TxnTypeList"), ("L", "EntityClassList"),
                      ("M", "ScopeList"), ("P", "TimelineList"),
                      ("E", "YesNoList")]:
        dv = DataValidation(type="list", formula1=name, allow_blank=True)
        wd.add_data_validation(dv)
        dv.add(f"{col}2:{col}{LAST_D}")

    # ---------------- Assignments
    hdr = ["Person", "Level (roster)", "Project ID", "Client", "Deal lifecycle",
           "Staffed as", "Active?", "Override hrs/wk", "Planned hrs/wk (used)",
           "Deal probability", "Weighted hrs/wk", "Deal start", "Deal end",
           "Eff. start (calc)", "Eff. end (calc)", "Source status (Jul 9)", "Notes"]
    style_header(wa, 1, hdr, [16, 16, 24, 26, 12, 16, 10, 10, 10, 10, 10, 11, 11,
                              11, 11, 24, 34])
    rows = [dict(person="Tania", code="EXAMPLE_0", staffed_as="Partner",
                 active="Inactive", raw="",
                 note="EXAMPLE row - Inactive so it counts nowhere; delete "
                      "any time")] + assign_rows
    r = 2
    for a in rows:
        wcell(wa, r, 1, a["person"], F_INPUT)
        wcell(wa, r, 3, a["code"], F_INPUT)
        wcell(wa, r, 6, a["staffed_as"], F_INPUT)
        wcell(wa, r, 7, a["active"], F_INPUT)
        if a["code"] == "EXAMPLE_0":
            wcell(wa, r, 8, 6, F_INPUT, fmt="0.0")
        wcell(wa, r, 16, a["raw"])
        wcell(wa, r, 17, a["note"], F_NOTE if "EXAMPLE" in a["note"] else F_BODY)
        r += 1
    for rr in range(2, LAST_A + 1):
        wcell(wa, rr, 2, f'=IF($A{rr}="","",IFERROR(INDEX(Roster!$B$2:$B${ROSTER_LAST},'
                         f'MATCH($A{rr},Roster!$A$2:$A${ROSTER_LAST},0)),"?"))', F_LINK)
        wcell(wa, rr, 4, safe_lookup(rr, LAST_D, f'Deals!$C$2:$C${LAST_D}', '"?"'),
              F_LINK)
        wcell(wa, rr, 5, f'=IF($C{rr}="","",IFERROR(INDEX(Deals!$F$2:$F${LAST_D},'
                         f'MATCH($C{rr},Deals!$A$2:$A${LAST_D},0)),"?"))', F_LINK)
        wcell(wa, rr, 9, f'=IF(OR($A{rr}="",$G{rr}<>"Active"),0,IF($H{rr}<>"",'
                         f'$H{rr},IFERROR(INDEX(Settings!$B$8:$B${7+len(DEFAULT_HOURS)},'
                         f'MATCH($F{rr},Settings!$A$8:$A${7+len(DEFAULT_HOURS)},0)),0)))',
              F_BODY, fmt="0.0")
        wcell(wa, rr, 10, f'=IF($C{rr}="",0,IFERROR(INDEX(Deals!$H$2:$H${LAST_D},'
                          f'MATCH($C{rr},Deals!$A$2:$A${LAST_D},0)),0))', F_LINK,
              fmt="0%")
        wcell(wa, rr, 11, f'=IF(Settings!$B$4="Yes",$I{rr}*$J{rr},$I{rr})',
              F_BODY, fmt="0.0")
        wcell(wa, rr, 12, safe_lookup(rr, LAST_D, f'Deals!$I$2:$I${LAST_D}'),
              F_LINK, fmt="yyyy-mm-dd")
        wcell(wa, rr, 13, safe_lookup(rr, LAST_D, f'Deals!$J$2:$J${LAST_D}'),
              F_LINK, fmt="yyyy-mm-dd")
        wcell(wa, rr, 14, f'=IF($L{rr}="",DATE(1900,1,1),$L{rr})', F_BODY,
              fmt="yyyy-mm-dd")
        wcell(wa, rr, 15, f'=IF($M{rr}="",DATE(2100,12,31),$M{rr})', F_BODY,
              fmt="yyyy-mm-dd")
        for ccol in range(1, 18):
            wa.cell(row=rr, column=ccol).border = THIN_BTM
    wa.freeze_panes = "D2"
    wa.conditional_formatting.add(
        f"A2:Q{LAST_A}",
        FormulaRule(formula=['$G2="Inactive"'], fill=FILL_GREY))
    for col, name in [("A", "RosterNames"), ("C", "DealCodes"),
                      ("F", "LevelList"), ("G", "ActiveList")]:
        dv = DataValidation(type="list", formula1=name, allow_blank=True)
        wa.add_data_validation(dv)
        dv.add(f"{col}2:{col}{LAST_A}")

    # ---------------- Capacity
    wcell(cap, 1, 1, "Capacity - planned hours per person per week", F_TITLE)
    wcell(cap, 2, 1, '="Probability weighting: "&Settings!$B$4&"   |   Deals '
                     'with no dates count in every week (fill Expected start/end '
                     'on the Deals tab to time-phase them)"', F_NOTE)
    wcell(cap, 4, 1, "Person", F_HDR, FILL_HDR)
    wcell(cap, 4, 2, "Cap hrs/wk", F_HDR, FILL_HDR)
    cap.column_dimensions["A"].width = 16
    cap.column_dimensions["B"].width = 10
    for w in range(N_WEEKS):
        col = 3 + w
        letter = get_column_letter(col)
        cap.column_dimensions[letter].width = 7
        f = "=Settings!$B$3" if w == 0 else f"={get_column_letter(col-1)}4+7"
        wcell(cap, 4, col, f, F_HDR, FILL_HDR, "dd-mmm",
              Alignment(horizontal="center"))
    first_p, last_p = 5, 5 + ROSTER_LAST - 2               # rows for roster 2..40
    for i, rr in enumerate(range(first_p, last_p + 1)):
        src = 2 + i
        wcell(cap, rr, 1, f'=IF(Roster!$A{src}="","",Roster!$A{src})', F_LINK)
        wcell(cap, rr, 2, f'=IF($A{rr}="","",IFERROR(INDEX(Roster!$D$2:$D${ROSTER_LAST},'
                          f'MATCH($A{rr},Roster!$A$2:$A${ROSTER_LAST},0)),""))', F_LINK,
              fmt="0")
        for w in range(N_WEEKS):
            col = 3 + w
            L = get_column_letter(col)
            f = (f'=IF($A{rr}="","",SUMIFS(Assignments!$K$2:$K${LAST_A},'
                 f'Assignments!$A$2:$A${LAST_A},$A{rr},'
                 f'Assignments!$G$2:$G${LAST_A},"Active",'
                 f'Assignments!$N$2:$N${LAST_A},"<="&{L}$4,'
                 f'Assignments!$O$2:$O${LAST_A},">="&{L}$4))')
            wcell(cap, rr, col, f, F_BODY, fmt="0.0;-0.0;")
    lastw = get_column_letter(2 + N_WEEKS)
    trow = last_p + 2
    wcell(cap, trow, 1, "Team planned", F_BOLD)
    wcell(cap, trow + 1, 1, "Team capacity", F_BOLD)
    wcell(cap, trow + 2, 1, "Headroom", F_BOLD)
    for w in range(N_WEEKS):
        L = get_column_letter(3 + w)
        wcell(cap, trow, 3 + w, f"=SUM({L}{first_p}:{L}{last_p})", F_BOLD, fmt="0")
        wcell(cap, trow + 1, 3 + w, f"=SUM(Roster!$D$2:$D${ROSTER_LAST})", F_BOLD,
              fmt="0")
        wcell(cap, trow + 2, 3 + w, f"={L}{trow+1}-{L}{trow}", F_BOLD, fmt="0")
    cap.freeze_panes = "C5"
    grid = f"C{first_p}:{lastw}{last_p}"
    cap.conditional_formatting.add(grid, FormulaRule(
        formula=[f"AND(ISNUMBER(C{first_p}),ISNUMBER($B{first_p}),"
                 f"C{first_p}>$B{first_p})"], fill=FILL_RED, stopIfTrue=True))
    cap.conditional_formatting.add(grid, FormulaRule(
        formula=[f"AND(ISNUMBER(C{first_p}),ISNUMBER($B{first_p}),"
                 f"C{first_p}>0.85*$B{first_p})"], fill=FILL_AMBER,
        stopIfTrue=True))
    cap.conditional_formatting.add(
        f"C{trow+2}:{lastw}{trow+2}",
        FormulaRule(formula=[f"C{trow+2}<0"], fill=FILL_RED))

    # ---------------- Review
    wcell(rv, 1, 1, "Migration review", F_TITLE)
    wcell(rv, 2, 1, "Built automatically from Scheduling_US_MA_July_9.xlsx. "
                    "Nothing was deleted. Section 1 needs a human answer (yellow "
                    "column); section 2 was resolved from the source data and is "
                    "here only as an audit trail.", F_NOTE)
    style_header(rv, 4, ["Topic", "Person", "Project", "Detail",
                         "Suggested action", "Your decision"],
                 [30, 14, 26, 60, 44, 24])

    # Items that genuinely need a person to decide (everything else is FYI).
    NEEDS_DECISION = {
        "Unclear personal status", "Ambiguous 'Kyle' rows skipped",
        "Name not in roster", "Possible duplicate engagement",
        "Engagement missing from register",
        "Proposal probabilities defaulted", "Deal dates blank",
        "Weekly capacity is a placeholder",
        "Level-default hours are placeholders",
    }
    blanket = [
        ("Proposal probabilities defaulted", "", "All 'Proposal' deals",
         "Probability prefilled at 50% (WIP deals at 100%)",
         "Set real probabilities in the yellow Probability cells (Deals tab)"),
        ("Deal dates blank", "", "All deals",
         "The July 9 file has no dates, so every deal counts in every week of "
         "the Capacity view",
         "Fill Expected start/end (yellow cells) to time-phase the load"),
        ("Weekly capacity is a placeholder", "", "All roster",
         "Everyone is set to 40 h/wk (Roster column D)",
         "Set each person's real weekly capacity"),
        ("Level-default hours are placeholders", "", "Settings B8:B14",
         "Hours/wk per level are directional guesses, not measured data",
         "Tune with the team at the first check-in"),
        ("Old 'Filtered Status' tab", "", "-",
         "It was a manual pivot of the per-person survey; superseded by the "
         "Assignments tab filters", "Nothing to do"),
    ]
    all_items = list(review) + blanket
    needs = sorted((it for it in all_items if it[0] in NEEDS_DECISION),
                   key=lambda x: x[0])
    resolved = sorted((it for it in all_items if it[0] not in NEEDS_DECISION),
                      key=lambda x: x[0])

    def decision_for(topic, proj):
        """The team's recorded answer for a shortlist item (2026-07-19 review)."""
        if topic == "Ambiguous 'Kyle' rows skipped":
            return DECISION_TEXT["kyle"], False
        if topic == "Deal dates blank":
            return DECISION_TEXT["dates"], False
        if topic == "Level-default hours are placeholders":
            return DECISION_TEXT["leveldefault"], True      # still partly open
        if topic == "Proposal probabilities defaulted":
            return DECISION_TEXT["prob"], True               # values still to come
        if topic == "Weekly capacity is a placeholder":
            return DECISION_TEXT["capacity"], False
        if topic == "Unclear personal status" and proj == "CLS_1":
            return DECISION_TEXT["cls"], False
        if topic == "Unclear personal status" and proj == "MIF_1":
            return DECISION_TEXT["mif"], False
        return None, True

    r = 5
    wcell(rv, r, 1, f"1) REVIEWED 2026-07-19  ({len(needs)} items)  -  team's "
                    f"answer in the last column; amber = still partly open",
          F_BOLD, FILL_YELLOW)
    r += 1
    for topic, person, proj, detail, action in needs:
        for ci, v in enumerate([topic, person, proj, detail, action], start=1):
            wcell(rv, r, ci, v, align=Alignment(wrap_text=True, vertical="top"))
        ans, still_open = decision_for(topic, proj)
        wcell(rv, r, 6, ans, F_INPUT if still_open else F_BODY,
              FILL_AMBER if still_open else None,
              align=Alignment(wrap_text=True, vertical="top"))
        r += 1
    r += 1
    wcell(rv, r, 1, f"2) RESOLVED FROM SOURCE DATA  ({len(resolved)} items)  -  "
                    f"audit trail, no action needed", F_BOLD, FILL_GREY)
    r += 1
    for topic, person, proj, detail, action in resolved:
        for ci, v in enumerate([topic, person, proj, detail, action], start=1):
            wcell(rv, r, ci, v, align=Alignment(wrap_text=True, vertical="top"))
        wcell(rv, r, 6, "resolved", F_NOTE)
        r += 1
    r += 2
    wcell(rv, r, 1, "Per-person status mapping (raw text -> Active?/Inactive)",
          F_BOLD)
    r += 1
    style_header(rv, r, ["Raw status (as typed)", "Mapped to", "", "", "", ""])
    r += 1
    for raw in sorted(STATUS_ACTIVE):
        wcell(rv, r, 1, raw)
        wcell(rv, r, 2, "Active")
        r += 1
    for raw in sorted(STATUS_INACTIVE | STATUS_REVIEW_INACTIVE):
        wcell(rv, r, 1, raw)
        wcell(rv, r, 2, "Inactive")
        r += 1
    rv.freeze_panes = "A5"

    # ---------------- Guide
    guide.column_dimensions["A"].width = 4
    guide.column_dimensions["B"].width = 120
    lines = [
        ("US M&A Tax - Scheduling & Capacity Tracker", F_TITLE),
        ("Built from Scheduling_US_MA_July_9.xlsx + the June 4 scheduling brief. "
         "This workbook is the bridge tracker: one source of truth for deals, "
         "staffing and weekly capacity.", F_NOTE),
        ("", F_BODY),
        ("TABS", F_BOLD),
        ("  Capacity - the dashboard: planned hours per person per week. Red = "
         "over capacity, amber = above 85%.", F_BODY),
        ("  Deals - one row per engagement: lifecycle, probability, dates and "
         "deal attributes (dropdowns).", F_BODY),
        ("  Assignments - one row per person on a deal. 'Active?' controls "
         "whether it counts. Hours come from the level default unless you set "
         "an Override. Columns N/O ('Eff. start/end') are calculation helpers "
         "that let Capacity add up fast - leave them alone, they just mirror "
         "columns L/M with a placeholder date when those are blank.", F_BODY),
        ("  Roster - the team, weekly capacity, live utilization.", F_BODY),
        ("  Settings - dropdown lists, level hour defaults, capacity window "
         "start, probability-weighting toggle.", F_BODY),
        ("  Review - section 1 lists what still needs a human answer (yellow "
         "'Your decision' column); section 2 is an audit trail of what was "
         "resolved automatically from the source data.", F_BODY),
        ("", F_BODY),
        ("COLOR CODE", F_BOLD),
        ("  Blue text = input you can edit    |    Black text = formula, leave "
         "alone    |    Green text = pulled from another tab", F_BODY),
        ("  Yellow fill = please fill in / review    |    Grey row = inactive "
         "(dead, closed, rolled off)", F_BODY),
        ("", F_BODY),
        ("HOW TO", F_BOLD),
        ("  Add a deal: next blank row on Deals - Project ID, client, Lifecycle, "
         "Probability, dates, attributes. Then staff it on Assignments.", F_BODY),
        ("  Staff someone: new row on Assignments - pick Person and Project ID, "
         "set Active? = Active. Hours default by level; type an Override for "
         "this deal if needed.", F_BODY),
        ("  Roll someone off: set their row's Active? to Inactive (keeps "
         "history - do not delete).", F_BODY),
        ("  Deal dies / closes: set Lifecycle on Deals; its assignments stop "
         "counting once you mark them Inactive (a grey row reminds you).",
         F_BODY),
        ("  New team member: add to Roster (name, level, capacity) - they "
         "appear in every dropdown.", F_BODY),
        ("  Weekly 15-min check-in: open Capacity (who is red/amber?), scan "
         "Deals for yellow probability/date cells to update, confirm staffing "
         "on incoming deals.", F_BODY),
        ("", F_BODY),
        ("ASSUMPTIONS BAKED IN (all editable)", F_BOLD),
        ("  1. Default hrs/wk by level (Settings B8:B14) are PLACEHOLDERS, not "
         "measured data - tune at the first check-in.", F_BODY),
        ("  2. Weekly capacity is 40h for everyone (Roster col D) - placeholder.",
         F_BODY),
        ("  3. WIP deals were prefilled at 100% probability, Proposals at 50% - "
         "review the yellow cells.", F_BODY),
        ("  4. Deals with blank dates count in every week of the Capacity view; "
         "fill dates to time-phase.", F_BODY),
        ("  5. Where the per-person survey said someone is done ('pencils "
         "down', 'closed', 'not involved'), that assignment was set Inactive - "
         "full mapping on the Review tab.", F_BODY),
        ("", F_BODY),
        ("Source: Scheduling_US_MA_July_9.xlsx (July 9) and Justine Morin's "
         "scheduling brief (June 4, 2026). Migration by scripts/"
         "build_workbook.py - regenerating overwrites manual edits, so edit "
         "here, not there.", F_NOTE),
    ]
    for i, (txt, font) in enumerate(lines, start=1):
        wcell(guide, i, 2, txt, font,
              align=Alignment(wrap_text=True, vertical="top"))
    guide.sheet_view.showGridLines = False

    # default font sweep for anything not explicitly styled
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for c in row:
                if c.value is not None and (c.font is None or
                                            c.font.name != ARIAL):
                    bold = c.font.bold if c.font else False
                    c.font = Font(name=ARIAL, size=10, bold=bold,
                                  color=c.font.color if c.font else None)

    wb.save(out_path)
    return dict(deals=len(deal_rows), assigns=len(assign_rows),
                review_needs=len(needs), review_resolved=len(resolved))


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "Scheduling_US_MA_July_9.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "US_MA_Tax_Scheduling_Tracker_v2.xlsx"
    deals, assigns, review, applied = extract(src)
    stats = build(deals, assigns, review, out)
    print(f"deals={stats['deals']} assignments={stats['assigns']} "
          f"statuses_applied={applied} "
          f"review_needs_decision={stats['review_needs']} "
          f"review_resolved={stats['review_resolved']}")
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
