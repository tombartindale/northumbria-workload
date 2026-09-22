#!/usr/bin/env python3
"""
generate_allocation_pdfs.py
===========================
Generate a one-page PDF workload-allocation summary for every member of staff
held in a School of Computer Science "Resources Planner" workbook.

One workbook contains the WHOLE roster: each member of staff is a single
column in the wide matrix on the 'Resources Planner' sheet (the first staff
column is 'Y'). The row layout is identical for everyone, so we read each
staff column in turn and emit one PDF per person.

Authoritative category totals (Teaching & Assessment / Leadership & Admin /
Research, Scholarship & Enterprise / Grand Total) are taken straight from the
sheet's own 'Summary / Totals' block. Individual line items (module codes,
semesters, roles and their hours) are read from the itemised sections above
that block. For each category we also emit a balancing line so the itemised
detail always reconciles to the official total — if that line is ever non-zero
it flags an extraction gap to fix.

Usage
-----
    python generate_allocation_pdfs.py INPUT.xlsm [-o OUTPUT_DIR] [--id NHQR5]

    INPUT.xlsm        the Resources Planner workbook (un-hidden master preferred)
    -o OUTPUT_DIR     where the PDFs are written (default: ./allocation_pdfs)
    --id ID           only generate for one staff ID (handy while iterating)

Requires: openpyxl, reportlab.  The workbook must have been saved by Excel so
that cached formula values are present (openpyxl reads cached values, it does
not recalculate). If values come back empty, open + save the file in Excel
first, or recalculate with LibreOffice.
"""

from __future__ import annotations
import argparse, os, re, sys, datetime
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle)
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF
from svglib.svglib import svg2rlg

# --------------------------------------------------------------------------
# Workbook layout constants  (edit here if the template ever changes)
# --------------------------------------------------------------------------
SHEET            = "Resources Planner"
ADDITIONAL_SHEET = "Additional Activities"
FIRST_STAFF_COL  = 25            # column 'Y' — first member of staff in the matrix

# Identity fields: row number -> key  (read from each staff member's own column)
IDENTITY_ROWS = {2: "status", 3: "forename", 4: "surname", 5: "fullname",
                 6: "id", 7: "subject", 8: "line_manager", 9: "position",
                 10: "fte", 11: "total_hours", 12: "allocated", 13: "remaining"}

# Body columns used when reading the itemised sections
COL_MARKER   = 1    # A  section headers live here
COL_B        = 2    # B  module code / role / partner
COL_TITLE    = 3    # C  module title / activity
COL_PERIOD   = 5    # E  teaching period (SEM1, SEM2, YL, ...)
COL_ACTIVITY = 11   # K  activity name / "Module Code (n)" / "Hours"
COL_TYPE     = 18   # R  row type for teaching grid: TEACH / ASSESS / TOTAL

# Section header text exactly as it appears in column A
SECTIONS = ["Teaching and Assessment", "Other Teaching Delivery",
            "Off Campus Delivery", "Projects & Supervision",
            "Programme Management", "Additional Roles / Activities",
            "Leadership", "Research, Scholarship and Enterprise",
            "Summary / Totals"]

# Labels in the Summary / Totals block (column J) that give the headline totals
SUMMARY_LABELS = {"teaching": "Total Teaching and Assessment",
                  "leadership": "Total Leadership and Admin",
                  "rse": "Total RSE Allocation",
                  "grand": "Grand Total",
                  "prep": "Preparation Time"}

CATEGORY_NAMES = {"teaching": "Teaching & Assessment",
                  "leadership": "Leadership & Admin",
                  "rse": "Research, Scholarship & Enterprise"}

# Compact labels for narrow table columns (avoids overflow)
CATEGORY_SHORT = {"teaching": "Teaching & Assessment",
                  "leadership": "Leadership & Admin",
                  "rse": "RSE"}

# Labels suppressed from the "Other roles & activities" table (universal items not worth listing)
SUPPRESS_LABELS = {
    "General Administration",
    "Specific Administrative Responsibilities",  # rollup of leadership items on Additional Activities tab
    "Other Teaching Activity",                   # rollup of teaching items on Additional Activities tab
}

# Friendly names for teaching period codes
PERIOD_LABELS = {"SEM1": "Semester 1", "SEM2": "Semester 2", "SEM3": "Semester 3",
                 "YL": "Year Long", "YLSEM3": "Semester 3",
                 "TP1": "Trimester 1", "TP2": "Trimester 2",
                 "TP3": "Trimester 3", "TP4": "Trimester 4", "TPYL": "Year Long"}

# Brand colours
TEAL   = colors.HexColor("#1F4E5F")
TEAL_L = colors.HexColor("#E8EEF0")
GREY   = colors.HexColor("#666666")
LINE   = colors.HexColor("#BBBBBB")

# Pie chart colour palette
PIE_PALETTE = [
    colors.HexColor("#1F4E5F"), colors.HexColor("#2E7D9B"),
    colors.HexColor("#3AABBF"), colors.HexColor("#2E6B4F"),
    colors.HexColor("#4F9E6B"), colors.HexColor("#CC9B4F"),
    colors.HexColor("#9B6B2E"), colors.HexColor("#7B2E6B"),
    colors.HexColor("#B36BA3"), colors.HexColor("#CC4F4F"),
]

# Category dot colours (match the first three pie palette entries)
CAT_DOT_COLORS = {
    "teaching":   PIE_PALETTE[0],
    "leadership": PIE_PALETTE[1],
    "rse":        PIE_PALETTE[2],
}

BANNER_H = 25 * mm   # full-bleed banner drawn outside the page margin frame

# For these modules the module-tutor hours represent a supervision load.
# tutor_hours / divisor = number of students supervised; shown in the tutor column.
# Teach/assess hours are real contact time and are displayed normally.
SUPERVISION_MODULES = {
    "KV6013": {"divisor": 12, "label": "UG students"},
    "KF7029": {"divisor":  8, "label": "PG students"},
    "KV7029": {"divisor":  8, "label": "PG students"},
    "PE7046": {"divisor":  6, "label": "PG students"},
}

# Modules where the team list is suppressed (too large to be useful)
NO_TEAM_MODULES = set(SUPERVISION_MODULES)

# Sort order for teaching periods
PERIOD_ORDER = {"SEM1": 1, "SEM2": 2, "SEM3": 3, "YL": 4, "YLSEM3": 5,
                "TP1": 6, "TP2": 7, "TP3": 8, "TP4": 9, "TPYL": 10}


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def num(v) -> float:
    """Coerce a cell value to float; blanks/text -> 0.0"""
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def period_label(code) -> str:
    if code in (None, ""):
        return ""
    return PERIOD_LABELS.get(str(code).strip(), str(code).strip())


def safe_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", (s or "").strip()).strip("_")


def clean_module_code(code) -> str:
    """Strip trailing annotations (e.g. 'KV7029 (sem 2)' -> 'KV7029') so
    module-tutor codes match the codes used in the Teaching and Assessment grid."""
    return re.sub(r"\s*\(.*\)\s*$", "", str(code or "").strip()).strip()


def parse_module_tutor_cell(raw):
    """Split a module-tutor code cell into (clean_code, period_hint).

    Staff are sometimes annotated as tutor for only part of a module's
    periods, e.g. 'KV7029 (sem 2)'. period_hint is a substring like 'SEM2'
    to match against a module row's period code, or None if the annotation
    doesn't look like a period (e.g. 'KV6014 (SHAPE)') or there is none.
    """
    raw = str(raw or "").strip()
    code = clean_module_code(raw)
    m = re.search(r"\((.*)\)\s*$", raw)
    period_hint = None
    if m:
        note = m.group(1).strip().upper().replace(" ", "")
        sem_m = re.fullmatch(r"SEM(\d)", note)
        if sem_m:
            period_hint = f"SEM{sem_m.group(1)}"
    return code, period_hint


def _period_matches(period_hint, period) -> bool:
    """True if a module row's period code belongs to the given period hint."""
    if period_hint is None:
        return True
    return period_hint in str(period or "").upper()


def find_sections(ws) -> dict:
    """Return {section_name: (start_row, end_row)} by scanning column A."""
    marks = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=COL_MARKER).value
        if isinstance(v, str) and v.strip() in SECTIONS:
            marks.append((r, v.strip()))
    out = {}
    for i, (r, name) in enumerate(marks):
        end = marks[i + 1][0] - 1 if i + 1 < len(marks) else ws.max_row
        out[name] = (r, end)
    return out


def summary_row(ws, sec, label, col=10) -> int | None:
    """Find a row in the Summary block whose column J (default) equals label."""
    start, end = sec["Summary / Totals"]
    for r in range(start, end + 1):
        if ws.cell(row=r, column=col).value == label:
            return r
    return None


def list_staff(ws):
    """Yield (col_index, staff_id, fullname) for every populated staff column."""
    for c in range(FIRST_STAFF_COL, ws.max_column + 1):
        sid = ws.cell(row=6, column=c).value
        full = ws.cell(row=5, column=c).value
        if sid and full:
            yield c, str(sid), str(full)


def build_module_tutors(ws, sec):
    """Return {module_code: [(name, period_hint), ...]} of who holds module-tutor
    hours for each module. period_hint is None (applies to all periods) or a
    period code like 'SEM2' parsed from an annotation such as 'KV7029 (sem 2)'."""
    staff_names = {col: f"{ws.cell(row=3, column=col).value or ''} {ws.cell(row=4, column=col).value or ''}".strip()
                   for col, _sid, _full in list_staff(ws)}
    tutors = defaultdict(set)
    pm_start, pm_end = sec["Programme Management"]
    for r in range(pm_start, pm_end + 1):
        lab = ws.cell(row=r, column=COL_ACTIVITY).value
        if isinstance(lab, str) and lab.startswith("Module Code"):
            for col, name in staff_names.items():
                if not name:
                    continue
                raw = ws.cell(row=r, column=col).value
                hrs = num(ws.cell(row=r + 1, column=col).value)
                if raw and hrs:
                    code, period_hint = parse_module_tutor_cell(raw)
                    tutors[code].add((name, period_hint))
    return {k: sorted(v) for k, v in tutors.items()}


def build_staff_emails(ws) -> dict:
    """Return {full_name: email} by matching staff IDs against staff_emails.csv."""
    import csv
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "staff_emails.csv")
    id_to_email = {}
    if os.path.exists(csv_path):
        with open(csv_path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("staff_id") and row.get("email"):
                    id_to_email[row["staff_id"].strip().upper()] = row["email"].strip()
    emails = {}
    for col, sid, _ in list_staff(ws):
        name = f"{ws.cell(row=3, column=col).value or ''} {ws.cell(row=4, column=col).value or ''}".strip()
        email = id_to_email.get(sid.strip().upper())
        if name and email:
            emails[name] = email
    return emails


def build_module_teams(ws, sec):
    """Return {(code, period): [sorted names]} excluding moderation-only contributors."""
    staff_names = {col: f"{ws.cell(row=3, column=col).value or ''} {ws.cell(row=4, column=col).value or ''}".strip()
                   for col, _sid, _full in list_staff(ws)}
    # Track non-moderation hours and moderation-only hours separately
    non_mod = defaultdict(set)   # (code, period) -> {names with teach/non-mod assess hours}
    mod_only = defaultdict(set)  # (code, period) -> {names with ONLY moderation hours}
    ta_start, ta_end = sec["Teaching and Assessment"]
    for r in range(ta_start + 2, ta_end + 1):
        code = ws.cell(row=r, column=COL_B).value
        rtype = ws.cell(row=r, column=COL_TYPE).value
        if not code or rtype not in ("TEACH", "ASSESS"):
            continue
        act = ws.cell(row=r, column=COL_ACTIVITY).value
        is_moderation = isinstance(act, str) and act.strip() == "Moderation"
        period = ws.cell(row=r, column=COL_PERIOD).value
        key = (str(code).strip(), period)
        for col, name in staff_names.items():
            if name and num(ws.cell(row=r, column=col).value) > 0:
                if is_moderation:
                    mod_only[key].add(name)
                else:
                    non_mod[key].add(name)
    # Only include staff who have non-moderation hours
    teams = {k: sorted(v) for k, v in non_mod.items()}
    return teams


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------
def extract_staff(ws, wadd, sec, col):
    """Build a structured allocation record for the staff member in `col`."""
    g = lambda r: ws.cell(row=r, column=col).value           # value in their column
    n = lambda r: num(ws.cell(row=r, column=col).value)      # numeric version

    ident = {k: g(r) for r, k in IDENTITY_ROWS.items()}

    # --- headline category totals (authoritative) ---------------------------
    totals = {}
    for key in ("teaching", "leadership", "rse", "grand"):
        r = summary_row(ws, sec, SUMMARY_LABELS[key])
        totals[key] = n(r) if r else 0.0
    prep_r = summary_row(ws, sec, SUMMARY_LABELS["prep"])
    prep_hours = n(prep_r) if prep_r else 0.0

    # --- module tutor hours by module code (Programme Management) -----------
    # entries: [(code, period_hint, hrs), ...]; period_hint is None (all
    # periods) or a period code like 'SEM2' parsed from an annotation.
    tutor_entries = []
    pm_start, pm_end = sec["Programme Management"]
    for r in range(pm_start, pm_end + 1):
        lab = ws.cell(row=r, column=COL_ACTIVITY).value
        if isinstance(lab, str) and lab.startswith("Module Code"):
            raw = g(r)
            hrs = n(r + 1)                # the row beneath holds the hours
            if raw and hrs:
                code, period_hint = parse_module_tutor_cell(raw)
                tutor_entries.append((code, period_hint, hrs))

    # --- teaching modules (Teaching and Assessment grid) --------------------
    ta_start, ta_end = sec["Teaching and Assessment"]
    modules = {}   # key -> dict
    for r in range(ta_start + 2, ta_end + 1):
        code = ws.cell(row=r, column=COL_B).value
        rtype = ws.cell(row=r, column=COL_TYPE).value
        if not code or rtype not in ("TEACH", "ASSESS", "TOTAL"):
            continue
        val = n(r)
        if val == 0:
            continue
        key = (str(code).strip(), ws.cell(row=r, column=COL_PERIOD).value)
        m = modules.setdefault(key, {"code": str(code).strip(),
                                     "title": ws.cell(row=r, column=COL_TITLE).value,
                                     "period": ws.cell(row=r, column=COL_PERIOD).value,
                                     "teach": 0.0, "assess": 0.0, "proj_super": 0.0})
        if rtype == "TEACH":
            m["teach"] += val
            act_label = ws.cell(row=r, column=COL_ACTIVITY).value
            if isinstance(act_label, str) and act_label.strip() == "Project Supervision":
                m["proj_super"] += val
        elif rtype == "ASSESS":
            act_label = ws.cell(row=r, column=COL_ACTIVITY).value
            if isinstance(act_label, str) and act_label.strip() == "Moderation":
                m.setdefault("moderation", 0.0)
                m["moderation"] += val
            else:
                m["assess"] += val
        # TOTAL rows are ignored for summing (used only as a sanity check)

    # attach module-tutor hours and build the final module list
    moderation_rows = []
    module_rows = []
    tutor_used = [False] * len(tutor_entries)
    for (code, period), m in sorted(modules.items()):
        m["tutor"] = 0.0
        for i, (t_code, t_hint, t_hrs) in enumerate(tutor_entries):
            if t_code == code and _period_matches(t_hint, period):
                m["tutor"] += t_hrs
                tutor_used[i] = True
        mod_hrs = m.pop("moderation", 0.0)
        if mod_hrs:
            moderation_rows.append({"code": code, "title": m["title"],
                                    "period": period, "hours": mod_hrs})
        m["total"] = m["teach"] + m["assess"] + m["tutor"]
        if m["total"] > 0:
            module_rows.append(m)
    # module-tutor entries with no matching teaching row (rare) -> their own line
    for i, (code, _hint, hrs) in enumerate(tutor_entries):
        if not tutor_used[i]:
            module_rows.append({"code": code, "title": "(module tutor only)",
                                "period": None, "teach": 0.0, "assess": 0.0,
                                "tutor": hrs, "total": hrs})

    # --- other itemised roles / activities ----------------------------------
    # category default per section; Programme Management & Additional Roles use
    # the model's own P-tag -> top-level mapping built from the Summary block.
    tag_cat = build_tag_category_map(ws, sec)
    others = []   # list of dicts: {label, category, hours}

    def add(label, category, hours):
        if hours and label not in SUPPRESS_LABELS:
            others.append({"label": label, "category": category, "hours": hours})

    # Other Teaching Delivery (activity in col B)
    s, e = sec["Other Teaching Delivery"]
    for r in range(s + 2, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        if lab and n(r):
            add(str(lab), "teaching", n(r))

    # Off Campus Delivery (partner carries forward in col B, activity in col C)
    s, e = sec["Off Campus Delivery"]
    partner = None
    for r in range(s + 2, e + 1):
        b = ws.cell(row=r, column=COL_B).value
        if b:
            partner = str(b)
        act = ws.cell(row=r, column=COL_TITLE).value
        if act and n(r) and not str(act).endswith("Total"):
            add(f"Off-campus: {partner} – {act}", "teaching", n(r))

    # Projects & Supervision (type in col B)
    _pgr_labels = {"PGR Supervision (1st Supervisor)", "PGR Supervision (2nd Supervisor)", "PGR Supervision (3rd Supervisor)"}
    _pgr_total = 0.0
    s, e = sec["Projects & Supervision"]
    for r in range(s + 2, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        if not lab or not n(r) or str(lab).endswith("Total"):
            continue
        if str(lab).strip() in _pgr_labels:
            _pgr_total += n(r)
        else:
            add(str(lab), "teaching", n(r))
    if _pgr_total:
        add("PGR Supervision", "teaching", _pgr_total)

    # Programme Management — non module-tutor roles (Programme Leader, Year Tutor…)
    s, e = sec["Programme Management"]
    for r in range(s + 2, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        actlab = ws.cell(row=r, column=COL_ACTIVITY).value
        if isinstance(actlab, str) and (actlab.startswith("Module Code")
                                        or actlab == "Hours"):
            continue   # handled as module-tutor
        if lab and n(r) and not str(lab).endswith("Total"):
            tag = ws.cell(row=r, column=16).value
            add(str(lab), tag_cat.get(tag, "leadership"), n(r))

    # Additional Roles / Activities
    s, e = sec["Additional Roles / Activities"]
    for r in range(s + 2, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        if lab and n(r) and not str(lab).endswith("Total"):
            tag = ws.cell(row=r, column=16).value
            add(str(lab), tag_cat.get(tag, "leadership"), n(r))

    # Leadership
    s, e = sec["Leadership"]
    for r in range(s + 2, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        if lab and n(r) and not str(lab).endswith("Total"):
            pct = ws.cell(row=r, column=5).value
            suffix = f" ({pct:.0%} FTE)" if isinstance(pct, (int, float)) and pct else ""
            add(f"{lab}{suffix}", "leadership", n(r))

    # Preparation time (only lives in the Summary block)
    add("Preparation time", "teaching", prep_hours)

    # Research, Scholarship and Enterprise — itemised from the section
    _rse_skip   = {"Research, Scholarship and Enterprise Activity", "Adjustment +/-"}
    _rse_rename = {"Research, Scholarship and Enterprise": "Ringfenced research time"}
    s, e = sec["Research, Scholarship and Enterprise"]
    for r in range(s + 1, e + 1):
        lab = ws.cell(row=r, column=COL_B).value
        if isinstance(lab, str) and lab.strip() and lab.strip() not in _rse_skip:
            hrs = n(r + 1)
            if hrs:
                add(_rse_rename.get(lab.strip(), lab.strip()), "rse", hrs)

    # --- Additional Activities tab ------------------------------------------
    if wadd is not None:
        for r in range(2, wadd.max_row + 1):
            row_id = wadd.cell(row=r, column=2).value          # Staff ID column
            if row_id and str(row_id).strip() == str(ident.get("id")).strip():
                activity = wadd.cell(row=r, column=4).value
                desc = wadd.cell(row=r, column=5).value
                hrs = num(wadd.cell(row=r, column=6).value)
                label = desc or activity or "unspecified"
                cat = "teaching" if isinstance(activity, str) and activity.strip() == "Other Teaching" else "leadership"
                add(str(label), cat, hrs)

    return {"ident": ident, "totals": totals, "modules": module_rows,
            "others": others, "moderation": moderation_rows}


def build_tag_category_map(ws, sec) -> dict:
    """Map each Summary-block P-tag to its top-level category by walking the
    Summary rows and tracking which header block (column D) we are inside."""
    start, end = sec["Summary / Totals"]
    headers = {"Teaching and Assessment": "teaching",
               "Leadership and Admin": "leadership",
               "Research, Scholarship and Enterprise": "rse"}
    current = "teaching"
    mapping = {}
    for r in range(start, end + 1):
        d = ws.cell(row=r, column=4).value
        if isinstance(d, str) and d.strip() in headers:
            current = headers[d.strip()]
        tag = ws.cell(row=r, column=16).value
        if tag and tag != "N/A":
            mapping.setdefault(tag, current)
    return mapping


# --------------------------------------------------------------------------
# Module-view extraction & PDF
# --------------------------------------------------------------------------

# Display order and short header labels for activity types in module PDFs
ACTIVITY_ORDER = [
    "Lectures", "Seminars", "Practicals", "Tutorials",
    "Demonstrations", "Supervised Work", "Project Supervision",
    "Marking", "Marking (Calculated)", "Moderation",
]
ACTIVITY_SHORT = {
    "Lectures":             "Lec",
    "Seminars":             "Sem",
    "Practicals":           "Prac",
    "Tutorials":            "Tut",
    "Demonstrations":       "Demo",
    "Supervised Work":      "Sup Work",
    "Project Supervision":  "Proj Sup",
    "Marking":              "Marking",
    "Marking (Calculated)": "Marking*",
    "Moderation":           "Mod",
}


COL_STUDENTS_EXPECTED = 12   # Column L — expected student intake (repurposed in Students Expected: TOTAL rows)

def extract_all_modules(ws, sec):
    """Return a dict keyed by (code, period) with title, student count, and per-staff activity hours."""
    staff_names = {
        col: f"{ws.cell(row=3, column=col).value or ''} "
             f"{ws.cell(row=4, column=col).value or ''}".strip()
        for col, _sid, _full in list_staff(ws)
    }
    ta_start, ta_end = sec["Teaching and Assessment"]
    modules = {}   # (code, period) -> {"title", "students", "staff"}
    for r in range(ta_start + 2, ta_end + 1):
        code = ws.cell(row=r, column=COL_B).value
        rtype = ws.cell(row=r, column=COL_TYPE).value
        if not code:
            continue
        period = ws.cell(row=r, column=COL_PERIOD).value
        key = (str(code).strip(), period)
        if key not in modules:
            modules[key] = {
                "title": ws.cell(row=r, column=COL_TITLE).value or "",
                "students": 0,
                "staff": defaultdict(lambda: defaultdict(float)),
            }
        act = ws.cell(row=r, column=COL_ACTIVITY).value
        if rtype == "TOTAL" and isinstance(act, str) and act.strip() == "Students Expected:":
            val = num(ws.cell(row=r, column=COL_STUDENTS_EXPECTED).value)
            if val:
                modules[key]["students"] = int(val)
            continue
        if rtype not in ("TEACH", "ASSESS"):
            continue
        if not isinstance(act, str) or not act.strip():
            continue
        act = act.strip()
        if act == "Marking (Calculated)":
            act = "Marking"
        for col, name in staff_names.items():
            if not name:
                continue
            val = num(ws.cell(row=r, column=col).value)
            if val:
                modules[key]["staff"][name][act] += val
    return modules


def build_module_pdf(code, title, periods_data, out_path, emails=None):
    """Render a PDF for a module with one section per period.

    periods_data: [(period, staff_dict)] sorted by PERIOD_ORDER
    staff_dict: {staff_name: {activity: hours}}
    """
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "NU_Logo_White.svg")
    logo_drawing = svg2rlg(logo_path) if os.path.exists(logo_path) else None
    if logo_drawing is not None:
        scale = (BANNER_H * 0.72) / logo_drawing.height
        logo_drawing.width *= scale
        logo_drawing.height *= scale
        logo_drawing.transform = (scale, 0, 0, scale, 0, 0)

    banner_title = f"{code}  {title}"
    snapshot_line = ("Static snapshot generated "
                     + datetime.date.today().strftime("%d %B %Y")
                     + ". Hours allocated to each member of staff.")

    def draw_banner(canvas, doc):
        canvas.saveState()
        pw, ph = A4
        canvas.setFillColor(colors.black)
        canvas.rect(0, ph - BANNER_H, pw, BANNER_H, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        mid = ph - BANNER_H / 2
        canvas.setFont("Helvetica-Bold", 13)
        canvas.drawString(6 * mm, mid + 4, banner_title)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(6 * mm, mid - 9, snapshot_line)
        if logo_drawing is not None:
            lx = pw - logo_drawing.width - 6 * mm
            ly = ph - BANNER_H + (BANNER_H - logo_drawing.height) / 2
            renderPDF.draw(logo_drawing, canvas, lx, ly)
        canvas.restoreState()

    MARGIN = 10 * mm
    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=BANNER_H + 6 * mm, bottomMargin=14 * mm,
                            title=f"Module Summary — {code}")
    DW = A4[0] - 2 * MARGIN
    H = ParagraphStyle
    GREY = colors.HexColor("#888888")
    LINE = colors.HexColor("#cccccc")
    sec_style = H("mod_sec", fontName="Helvetica-Bold", fontSize=11,
                  textColor=TEAL, leading=14)
    meta_style = H("mod_meta", fontName="Helvetica-Oblique", fontSize=8,
                   textColor=GREY, leading=11)
    disc_style = H("mod_disc", fontName="Helvetica-Oblique", fontSize=7.5,
                   textColor=GREY, leading=10)

    story = []
    first = True

    for period, period_info in periods_data:
        staff_data = period_info["staff"]
        students = period_info["students"]
        if not first:
            story.append(Spacer(1, 14))
        first = False

        period_label = PERIOD_LABELS.get(str(period).strip() if period else "", period or "")
        heading = period_label
        if students:
            heading += f"  —  {students} students expected"
        story.append(Paragraph(heading, sec_style))
        story.append(Spacer(1, 5))

        # Collect activities present in this period (exclude Moderation — shown separately)
        all_acts = []
        moderators = []
        for name, acts in staff_data.items():
            for a in acts:
                if a == "Moderation":
                    if name not in moderators:
                        moderators.append(name)
                elif a not in all_acts:
                    all_acts.append(a)
        all_acts.sort(key=lambda a: (ACTIVITY_ORDER.index(a) if a in ACTIVITY_ORDER else 99, a))

        short_hdrs = [ACTIVITY_SHORT.get(a, a) for a in all_acts]
        col_w_name = 0.38 * DW
        n_num = len(all_acts) + 1  # activity cols + Total
        col_w_each = (DW - col_w_name) / n_num
        col_widths = [col_w_name] + [col_w_each] * n_num

        rows = [["Staff Member"] + short_hdrs + ["Total"]]
        col_totals = defaultdict(float)
        for name, acts in sorted(staff_data.items()):
            non_mod = {a: h for a, h in acts.items() if a != "Moderation"}
            if not non_mod:
                continue
            row_total = sum(non_mod.values())
            rows.append([name] + [fmt(non_mod.get(a, 0.0)) for a in all_acts] + [fmt(row_total)])
            for a in all_acts:
                col_totals[a] += non_mod.get(a, 0.0)

        grand_total = sum(col_totals.values())
        total_row_idx = len(rows)
        rows.append(["Total"] + [fmt(col_totals[a]) for a in all_acts] + [fmt(grand_total)])

        tbl = Table(rows, colWidths=col_widths)
        st = [("BACKGROUND",    (0, 0), (-1, 0), TEAL),
              ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
              ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
              ("FONTSIZE",      (0, 0), (-1, -1), 8.5),
              ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
              ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
              ("GRID",          (0, 0), (-1, -1), 0.4, LINE),
              ("ROWBACKGROUNDS",(0, 1), (-1, -2), [colors.white, TEAL_L]),
              ("TOPPADDING",    (0, 0), (-1, -1), 3),
              ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
              ("LEFTPADDING",   (0, 0), (-1, -1), 5),
              ("FONTNAME",      (0, total_row_idx), (-1, total_row_idx), "Helvetica-Bold"),
              ("BACKGROUND",    (0, total_row_idx), (-1, total_row_idx), TEAL_L)]
        tbl.setStyle(TableStyle(st))
        story.append(tbl)

        if moderators:
            story.append(Spacer(1, 4))
            story.append(Paragraph(
                f"<b>Moderator:</b> {', '.join(sorted(moderators))}", meta_style))

        if emails:
            teaching_names = sorted(
                name for name, acts in staff_data.items()
                if any(a != "Moderation" for a in acts)
            )
            team_emails = [emails[n] for n in teaching_names if n in emails]
            if team_emails:
                story.append(Spacer(1, 4))
                story.append(Paragraph(
                    f"<b>Teaching team emails (copy into Outlook):</b> {'; '.join(team_emails)}",
                    meta_style))

    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "DRAFT — FOR DISCUSSION ONLY. This document is an automatically generated, "
        "prototype summary of workload data held in the School Resources Planner. "
        "Figures are indicative only and do not constitute a formal workload agreement. "
        "Please contact your Subject Head if you have any questions.", disc_style))

    doc.build(story, onFirstPage=draw_banner, onLaterPages=draw_banner)


# --------------------------------------------------------------------------
# PDF rendering helpers
# --------------------------------------------------------------------------
def _pie_drawing(items, title, w, h):
    """Pie chart with a right-hand legend. items: [(label, value), ...] all > 0."""
    d = Drawing(w, h)
    total = sum(v for _, v in items)
    if not total:
        return d

    d.add(String(w / 2, h - 10, title, fontName="Helvetica-Bold", fontSize=8,
                 textAnchor="middle", fillColor=TEAL))

    pie_diam = min(w * 0.46, h - 22)
    pie = Pie()
    pie.x = 4
    pie.y = (h - 16 - pie_diam)
    pie.width = pie_diam
    pie.height = pie_diam
    pie.data = [v for _, v in items]
    for i in range(len(items)):
        pie.slices[i].fillColor = PIE_PALETTE[i % len(PIE_PALETTE)]
        pie.slices[i].strokeColor = colors.white
        pie.slices[i].strokeWidth = 0.5
        pie.slices[i].label_visible = False
    d.add(pie)

    # legend to the right
    leg_x = pie.x + pie_diam + 8
    leg_top = pie.y + pie_diam
    row_h = pie_diam / len(items)
    box = 7
    max_chars = max(8, int((w - leg_x - box - 12) / 4.1))
    for i, (label, val) in enumerate(items):
        pct = 100 * val / total
        y = leg_top - i * row_h - row_h / 2
        d.add(Rect(leg_x, y - box / 2, box, box,
                   fillColor=PIE_PALETTE[i % len(PIE_PALETTE)], strokeColor=None))
        short = label[:max_chars] + "…" if len(label) > max_chars else label
        d.add(String(leg_x + box + 3, y - 3,
                     f"{short} ({pct:.0f}%)",
                     fontName="Helvetica", fontSize=6.5, fillColor=colors.black))
    return d


def fmt(x) -> str:
    if x in (None, ""):
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    if not isinstance(x, (int, float)):
        return str(x)
    return f"{x:,}" if isinstance(x, (int,)) else f"{x:,.2f}"


def build_pdf(data, out_path, teams=None, tutors=None):
    ident, totals = data["ident"], data["totals"]
    styles = getSampleStyleSheet()
    H = lambda **kw: ParagraphStyle("h", parent=styles["Normal"], **kw)
    name_style  = H(fontName="Helvetica-Bold", fontSize=22, textColor=colors.black, leading=26)
    sub_style   = H(fontName="Helvetica", fontSize=9, textColor=GREY, leading=12)
    sec_style   = H(fontName="Helvetica-Bold", fontSize=11, textColor=TEAL,
                    spaceBefore=10, spaceAfter=4)
    cell        = H(fontName="Helvetica", fontSize=9, leading=11)
    cell_b      = H(fontName="Helvetica-Bold", fontSize=9, leading=11)

    # ---- Logo (scaled once for use in the canvas callback) ------------
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "NU_Logo_White.svg")
    logo_drawing = svg2rlg(logo_path) if os.path.exists(logo_path) else None
    if logo_drawing is not None:
        scale = (BANNER_H * 0.72) / logo_drawing.height
        logo_drawing.width  *= scale
        logo_drawing.height *= scale
        logo_drawing.transform = (scale, 0, 0, scale, 0, 0)

    snapshot_line = ("Static snapshot generated "
                     + datetime.date.today().strftime("%d %B %Y")
                     + ". Hours allocated to this member of staff only.")

    def draw_banner(canvas, doc):
        canvas.saveState()
        pw, ph = A4
        canvas.setFillColor(colors.black)
        canvas.rect(0, ph - BANNER_H, pw, BANNER_H, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        mid = ph - BANNER_H / 2
        canvas.setFont("Helvetica-Bold", 13)
        canvas.drawString(6 * mm, mid + 4, "Workload draft")
        canvas.setFont("Helvetica", 8)
        canvas.drawString(6 * mm, mid - 9, snapshot_line)
        if logo_drawing is not None:
            lx = pw - logo_drawing.width - 6 * mm
            ly = ph - BANNER_H + (BANNER_H - logo_drawing.height) / 2
            renderPDF.draw(logo_drawing, canvas, lx, ly)
        canvas.restoreState()

    MARGIN = 10 * mm
    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=BANNER_H + 6*mm, bottomMargin=14*mm,
                            title=f"Workload Allocation — {ident.get('fullname')}")
    DW = A4[0] - 2 * MARGIN   # usable page width for full-width tables
    story = []

    # ---- Name & identity ----------------------------------------------
    story.append(Paragraph(
        f"{ident.get('forename') or ''} {ident.get('surname') or ''}".strip(), name_style))
    story.append(Spacer(1, 6))

    # ---- Pie charts ---------------------------------------------------
    cat_items = [(CATEGORY_NAMES[k], totals[k])
                 for k in ("teaching", "leadership", "rse") if totals[k]]
    mod_totals = defaultdict(float)
    for m in data["modules"]:
        contact = m["teach"] + m["assess"]
        if contact:
            mod_totals[m["code"]] += contact
    mod_items = sorted(((c, h) for c, h in mod_totals.items() if h), key=lambda x: -x[1])

    if cat_items or mod_items:
        cw = DW / 2
        ph = 150
        pie_row = [
            _pie_drawing(cat_items, "Category breakdown", cw, ph) if cat_items
            else Spacer(cw, ph),
            _pie_drawing(mod_items, "Hours by module", cw, ph) if mod_items
            else Spacer(cw, ph),
        ]
        pt = Table([pie_row], colWidths=[cw, cw], spaceAfter=4)
        pt.setStyle(TableStyle([("VALIGN",        (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING",   (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
                                ("TOPPADDING",    (0, 0), (-1, -1), 0),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        story.append(pt)
        story.append(Spacer(1, 4))

    # ---- Workload summary by category ---------------------------------
    story.append(Paragraph("Workload summary", sec_style))
    alloc = totals["grand"] or 1
    rows = [["Category", "Hours", "% of allocated"]]
    for key in ("teaching", "leadership", "rse"):
        rows.append([CATEGORY_NAMES[key], fmt(totals[key]),
                     f"{100*totals[key]/alloc:.1f}%"])
    rows.append(["Grand total", fmt(totals["grand"]), "100.0%"])
    rows.append(["Remaining capacity", fmt(ident.get("remaining")), ""])
    summ = Table(rows, colWidths=[0.66*DW, 0.17*DW, 0.17*DW])
    summ.setStyle(_table_style(total_rows=[len(rows)-2], faint_rows=[len(rows)-1]))
    story.append(summ)

    # ---- Teaching modules — grouped by semester -----------------------
    if data["modules"]:
        story.append(Paragraph("Teaching modules", sec_style))
        sorted_mods = sorted(
            data["modules"],
            key=lambda m: (PERIOD_ORDER.get(str(m["period"] or "").strip(), 50), m["code"]))
        my_name = f"{ident.get('forename') or ''} {ident.get('surname') or ''}".strip()
        rows = [["Module", "Title", "Teaching", "Assessment", "Module tutor", "Total"]]
        span_rows, prev_period = [], object()
        for m in sorted_mods:
            p = m["period"]
            if p != prev_period:
                rows.append([period_label(p) if p else "Other", "", "", "", "", ""])
                span_rows.append(len(rows) - 1)
                prev_period = p

            tutor_names = {name for name, hint in (tutors or {}).get(m["code"], [])
                           if _period_matches(hint, m["period"])}
            sup = SUPERVISION_MODULES.get(m["code"])
            if m["code"] in NO_TEAM_MODULES:
                # Team suppressed — just show the module tutor
                if my_name in tutor_names:
                    sub_line = "<b>You</b>"
                elif tutor_names:
                    sub_line = ", ".join(f"<b>{t}</b> (MT)" for t in sorted(tutor_names))
                else:
                    sub_line = None
            else:
                # Regular module: tutor(s) bold first, then remaining team
                team = [n for n in (teams or {}).get((m["code"], m["period"]), [])
                        if n != my_name]
                if my_name in tutor_names:
                    tutor_part = "<b>You</b>"
                elif tutor_names:
                    tutor_part = ", ".join(f"<b>{t}</b> (MT)" for t in sorted(tutor_names))
                else:
                    tutor_part = None
                others = [n for n in team if n not in tutor_names]
                parts = ([tutor_part] if tutor_part else []) + others
                sub_line = ", ".join(parts) if parts else None

            if sub_line:
                title_cell = Paragraph(
                    f'{_short(m["title"], 42)}<br/>'
                    f'<font size="7" color="#888888">{sub_line}</font>',
                    cell)
            else:
                title_cell = _short(m["title"], 42)
            proj_super = m.get("proj_super", 0.0)
            if sup and proj_super:
                n = int(round(proj_super / sup["divisor"]))
                tutor_cell = f"{n} {sup['label']}"
            else:
                tutor_cell = fmt(m["tutor"])
            code_cell = Paragraph(m["code"], cell) if " " in (m["code"] or "") else (m["code"] or "")
            rows.append([code_cell, title_cell,
                         fmt(m["teach"]), fmt(m["assess"]),
                         tutor_cell, fmt(m["total"])])
        tbl = Table(rows, colWidths=[0.10*DW, 0.45*DW, 0.10*DW, 0.12*DW, 0.12*DW, 0.11*DW])
        tbl.setStyle(_table_style(numeric_from_col=2, span_rows=span_rows))
        story.append(tbl)

    # ---- Moderation — one line per module -------------------------------
    if data["moderation"]:
        story.append(Paragraph("Moderation", sec_style))
        sorted_mod = sorted(
            data["moderation"],
            key=lambda m: (PERIOD_ORDER.get(str(m["period"] or "").strip(), 50), m["code"]))
        rows = [["Module", "Hours"]]
        for m in sorted_mod:
            label = f"{period_label(m['period']) if m['period'] else 'Other'} – {m['code']} {_short(m['title'], 45)}"
            rows.append([label, fmt(m["hours"])])
        tbl = Table(rows, colWidths=[0.87*DW, 0.13*DW])
        tbl.setStyle(_table_style(numeric_from_col=1))
        story.append(tbl)

    # ---- Other roles & activities -------------------------------------
    story.append(Paragraph("Other roles & activities", sec_style))
    DOT_W = 4 * mm
    rows = [["", "Activity", "Hours"]]
    dot_rows = []
    item_sum = {"teaching": 0.0, "leadership": 0.0, "rse": 0.0}
    module_teaching = sum(m["teach"] + m["assess"] + m["tutor"] for m in data["modules"])
    item_sum["teaching"] += module_teaching
    item_sum["teaching"] += sum(m["hours"] for m in data["moderation"])
    for o in sorted(data["others"], key=lambda x: (x["category"], -x["hours"])):
        dot_rows.append((len(rows), CAT_DOT_COLORS.get(o["category"], GREY)))
        rows.append(["", _short(o["label"], 72), fmt(o["hours"])])
        item_sum[o["category"]] += o["hours"]
    gap_labels = {"teaching": f"Other / unallocated within {CATEGORY_NAMES['teaching']}",
                  "leadership": "General Administration",
                  "rse": f"Other / unallocated within {CATEGORY_NAMES['rse']}"}
    for key in ("teaching", "leadership", "rse"):
        gap = round(totals[key] - item_sum[key], 2)
        if gap >= 0.5:
            dot_rows.append((len(rows), CAT_DOT_COLORS.get(key, GREY)))
            rows.append(["", gap_labels[key], fmt(gap)])
    tbl = Table(rows, colWidths=[DOT_W, DW - DOT_W - 0.13*DW, 0.13*DW])
    tbl.setStyle(_table_style(numeric_from_col=2, dot_rows=dot_rows))
    story.append(tbl)

    disclaimer_style = H(fontName="Helvetica-Oblique", fontSize=7.5, textColor=GREY,
                         leading=10, spaceBefore=12,
                         borderPad=4, borderColor=LINE, borderWidth=0.4,
                         borderRadius=2)
    story.append(Paragraph(
        "DRAFT — FOR DISCUSSION ONLY. This document is an automatically generated, "
        "prototype summary of workload data held in the School Resources Planner. "
        "Figures are indicative only and do not constitute a formal workload agreement. "
        "Please contact your Subject Head if you have any questions.",
        disclaimer_style))

    doc.build(story, onFirstPage=draw_banner, onLaterPages=draw_banner)


def _short(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[: n - 1] + "…"


def _table_style(numeric_from_col=1, total_rows=None, faint_rows=None, span_rows=None,
                 dot_rows=None):
    total_rows = total_rows or []
    faint_rows = faint_rows or []
    span_rows  = span_rows  or []
    dot_rows   = dot_rows   or []   # list of (row_index, color)
    st = [("BACKGROUND", (0, 0), (-1, 0), TEAL),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
          ("FONTSIZE", (0, 0), (-1, -1), 8.5),
          ("ALIGN", (numeric_from_col, 1), (-1, -1), "RIGHT"),
          ("ALIGN", (0, 0), (numeric_from_col-1, 0), "LEFT"),
          ("GRID", (0, 0), (-1, -1), 0.4, LINE),
          ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, TEAL_L]),
          ("TOPPADDING", (0, 0), (-1, -1), 3),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
          ("LEFTPADDING", (0, 0), (-1, -1), 5)]
    for r in total_rows:
        st += [("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"),
               ("BACKGROUND", (0, r), (-1, r), TEAL_L)]
    for r in faint_rows:
        st += [("TEXTCOLOR", (0, r), (-1, r), GREY)]
    for r in span_rows:
        st += [("SPAN", (0, r), (-1, r)),
               ("BACKGROUND", (0, r), (-1, r), TEAL_L),
               ("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"),
               ("TEXTCOLOR", (0, r), (-1, r), TEAL),
               ("ALIGN", (0, r), (-1, r), "LEFT")]
    for r, dot_color in dot_rows:
        st += [("BACKGROUND",    (0, r), (0, r), dot_color),
               ("LEFTPADDING",   (0, r), (0, r), 0),
               ("RIGHTPADDING",  (0, r), (0, r), 0),
               ("TOPPADDING",    (0, r), (0, r), 0),
               ("BOTTOMPADDING", (0, r), (0, r), 0)]
    return TableStyle(st)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Generate per-staff allocation PDFs.")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--id", default=None, help="only this staff ID")
    ap.add_argument("--skip-empty", action="store_true",
                    help="skip staff whose grand total is zero")
    ap.add_argument("--modules", action="store_true",
                    help="generate per-module PDFs instead of per-staff PDFs")
    ap.add_argument("--module", default=None,
                    help="only generate for one module code (use with --modules)")
    args = ap.parse_args()

    if args.out is None:
        args.out = "module_pdfs" if args.modules else "allocation_pdfs"
    os.makedirs(args.out, exist_ok=True)
    wb = openpyxl.load_workbook(args.workbook, data_only=True)
    ws = wb[SHEET]
    wadd = wb[ADDITIONAL_SHEET] if ADDITIONAL_SHEET in wb.sheetnames else None
    sec = find_sections(ws)

    if args.modules:
        print("Extracting module data…")
        staff_emails = build_staff_emails(ws)
        all_modules = extract_all_modules(ws, sec)
        # Group by code
        by_code = defaultdict(dict)
        titles = {}
        for (code, period), info in all_modules.items():
            if info["staff"]:
                by_code[code][period] = {"staff": info["staff"], "students": info["students"]}
                titles[code] = info["title"]
        made = 0
        for code in sorted(by_code):
            if args.module and code != args.module:
                continue
            periods_data = sorted(
                by_code[code].items(),
                key=lambda x: PERIOD_ORDER.get(str(x[0]).strip() if x[0] else "", 99)
            )
            fname = f"{safe_name(code)}_Module.pdf"
            out_path = os.path.join(args.out, fname)
            build_module_pdf(code, titles[code], periods_data, out_path, emails=staff_emails)
            made += 1
            print(f"  {code:<10}  ->  {fname}")
        print(f"\nDone. {made} module PDF(s) written to '{args.out}'.")
        return

    print("Building module team and tutor maps…")
    teams  = build_module_teams(ws, sec)
    tutors = build_module_tutors(ws, sec)

    made = skipped = 0
    for col, sid, full in list_staff(ws):
        if args.id and sid != args.id:
            continue
        data = extract_staff(ws, wadd, sec, col)
        if args.skip_empty and not data["totals"]["grand"]:
            skipped += 1
            continue
        surname = safe_name(data["ident"].get("surname") or "")
        forename = safe_name(data["ident"].get("forename") or "")
        fname = f"{surname}_{forename}_{safe_name(sid)}_Allocation.pdf"
        build_pdf(data, os.path.join(args.out, fname), teams, tutors)
        made += 1
        print(f"  {get_column_letter(col):>3}  {full:<32}  ->  {fname}")

    print(f"\nDone. {made} PDF(s) written to '{args.out}'."
          + (f" {skipped} empty staff skipped." if args.skip_empty else ""))


if __name__ == "__main__":
    main()
