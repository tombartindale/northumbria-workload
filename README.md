# Workload Allocation PDF Generator

Generates a one-page PDF workload summary for **every member of staff** held in a
School of Computer Science *Resources Planner* workbook (`.xlsm`/`.xlsx`).

This document explains how the source workbook is structured, exactly how the
script reads it, and where to change things — it is written so the script can be
handed to **Claude Code** for further refinement of the output.

---

## 1. What you get

For each member of staff the script writes `Surname_Forename_ID_Allocation.pdf`
containing:

1. **Identity block** — name, ID, position, subject area, line manager, FTE,
   total/allocated/remaining hours.
2. **Workload summary** — the three headline categories (Teaching & Assessment,
   Leadership & Admin, Research/Scholarship/Enterprise), each with hours and % of
   the allocated total, plus the grand total and remaining capacity.
3. **Teaching modules** — one row per allocated module: module code, title,
   semester, and a breakdown of teaching / assessment / module-tutor hours.
4. **Other roles & activities** — every other allocated item (leadership roles,
   administration, supervision, off-campus delivery, preparation time, RSE,
   anything on the *Additional Activities* tab), each with its category and hours.

Run it once against a single workbook and it produces the whole department.

---

## 2. The key fact about the source workbook

**One workbook already contains the entire roster.** On the `Resources Planner`
sheet there is a wide matrix in which **each member of staff is a single column**,
starting at column **`Y`** (column index 25). The per-person files you have in your
directory are just filtered *views* of this same matrix — so you do **not** need to
open 90+ files. Point the script at one workbook (ideally the **un-hidden master**,
because every cached value is guaranteed present) and it reads each staff column in
turn.

### Layout of a staff column

The same rows mean the same thing in every staff column:

| Row | Meaning            | Row | Meaning           |
|-----|--------------------|-----|-------------------|
| 2   | Status             | 8   | Line manager      |
| 3   | Forename           | 9   | Position          |
| 4   | Surname            | 10  | FTE               |
| 5   | Full name          | 11  | Total hours       |
| 6   | Staff ID           | 12  | Allocated hours   |
| 7   | Subject area       | 13  | Remaining hours   |

### Sections (down column A)

The body is divided into labelled sections, each a header in column **A**:

```
Teaching and Assessment
Other Teaching Delivery
Off Campus Delivery
Projects & Supervision
Programme Management
Additional Roles / Activities
Leadership
Research, Scholarship and Enterprise
Summary / Totals
```

The script finds these dynamically, so inserting rows inside a section will **not**
break it.

### The two numbers people confuse

In the *Teaching and Assessment* grid almost every module has a non-zero
**Activity Hours** value — that is the **whole school's** teaching demand for that
module, *not* this person's share. The hours actually given to a member of staff
live in **their own matrix column**. The script therefore reads each person's
column and ignores Activity Hours entirely. Filtering on Activity Hours instead
would wrongly pull back the entire module catalogue.

### The Summary / Totals block is authoritative

Rows under *Summary / Totals* roll everything up into the three official
categories and a grand total (which equals the staff member's `Allocated hours`).
The script takes the **headline category totals from here** and treats them as
ground truth. It then itemises the detail from the sections above and, per
category, prints a **balancing line** if the items don't add up to the official
total. Across the test workbook all 76 active staff reconciled with **zero**
balancing lines — but the mechanism means the PDF can never silently mis-state a
total.

---

## 3. How the script works (pipeline)

`generate_allocation_pdfs.py`

1. **`find_sections(ws)`** — scans column A for the known section headers and
   returns `{name: (start_row, end_row)}`.
2. **`list_staff(ws)`** — yields every populated staff column (`col, id, fullname`)
   from column `Y` rightwards.
3. **`build_tag_category_map(ws, sec)`** — walks the *Summary / Totals* block to
   learn which top-level category each internal tag (`MOD_TUTOR`, `PROG_LEAD`,
   `GEN_ADM`, `LEADERSHIP`, `RESEARCH`, …) belongs to. Used so that, e.g.,
   *Programme Leader* is filed under Leadership & Admin while *Module Tutor* is
   filed under Teaching & Assessment — using the workbook's own logic rather than
   a hard-coded guess.
4. **`extract_staff(ws, wadd, sec, col)`** — the heart of it. For one staff column:
   - reads identity + the four headline totals;
   - builds **teaching modules** by grouping the T&A grid on `(module code,
     period)` and summing the staff column for `TEACH` vs `ASSESS` rows (`TOTAL`
     rows are used only as a sanity check);
   - reads **module-tutor** hours from *Programme Management* (each
     `Module Code (n)` cell with the hours in the row beneath) and attaches them to
     the matching module;
   - collects **other items** from every remaining section, plus *Preparation
     Time* (which only exists in the Summary block), plus any rows on the
     *Additional Activities* tab matching this Staff ID;
   - returns a tidy dict.
5. **`build_pdf(data, out_path)`** — renders the dict with ReportLab Platypus
   (identity table, summary table, modules table, other-roles table, footer).

### Reconciliation guarantee

`build_pdf` sums the itemised lines per category and compares to the authoritative
totals. Any difference ≥ 0.5 h appears as
`Other / unallocated within <category>`. **If you ever see that line, it means a
section/role the extractor doesn't yet itemise — that's the place to extend
`extract_staff`.**

---

## 4. Running it

```bash
pip install openpyxl reportlab          # one-off

# whole department
python generate_allocation_pdfs.py MASTER.xlsm -o allocation_pdfs --skip-empty

# a single person while iterating on layout
python generate_allocation_pdfs.py MASTER.xlsm --id NHQR5 -o preview
```

| Flag           | Effect                                                        |
|----------------|---------------------------------------------------------------|
| `-o DIR`       | output folder (default `allocation_pdfs`)                     |
| `--id ID`      | only generate for one staff ID                                |
| `--skip-empty` | skip staff whose grand total is 0 (not yet allocated)         |

### Requirement: cached values must be present

`openpyxl` reads **cached** formula results — it does not recalculate. The master
workbook must have been **saved by Excel** so the cache is current. If figures come
back as 0/blank, open and re-save in Excel, or recalculate with LibreOffice:

```bash
soffice --headless --convert-to xlsx --calc MASTER.xlsm   # forces a recalc
```

---

## 5. Things you may want to refine (hooks for Claude Code)

Everything below is intentionally isolated so it's easy to change:

- **Template constants** (top of file): `FIRST_STAFF_COL`, `IDENTITY_ROWS`,
  `SECTIONS`, `SUMMARY_LABELS`, `COL_*`. Update these if the spreadsheet template
  changes (new section, moved column, etc.).
- **Branding & layout**: colours (`TEAL`, `TEAL_L`), the `_table_style` helper, and
  the table `colWidths` in `build_pdf`. Swap in a logo by adding an `Image` flowable
  at the top of `story`.
- **Period names**: `PERIOD_LABELS` maps codes (`SEM1`, `YL`, `TP2`…) to friendly
  text. Add any missing codes here.
- **Category assignment**: handled by `build_tag_category_map` + section defaults.
  If a role lands in the wrong bucket, check its tag in the Summary block.
- **Module grouping**: currently `(code, period)`. If you'd rather collapse all
  occurrences of a module into one row, change the `key` in `extract_staff`.
- **Page size / one-vs-multi page**: `SimpleDocTemplate(... pagesize=A4 ...)`. The
  content auto-flows to a second page if a staff member has many modules.
- **Output naming**: `fname` in `main()`.

### Possible enhancements
- A cover index page listing everyone and their grand total.
- A small horizontal bar chart of the three-way category split.
- Combine all PDFs into one book with `pypdf` (merge step after the loop).
- Emit a CSV/Excel audit alongside the PDFs for QA.

---

## 6. Worked example (Bartindale, Tom — NHQR5)

| Category                         | Hours    | %     |
|----------------------------------|---------:|------:|
| Teaching & Assessment            | 369.00   | 23.5% |
| Leadership & Admin               | 713.00   | 45.3% |
| Research, Scholarship & Enterprise | 491.04 | 31.2% |
| **Grand total**                  | **1,573.04** | 100% |

Teaching modules: **KV6013** Computing Project (Year Long) — 16 teaching / 2
assessment / 60 module-tutor; **KV7011** Multimodal Interfaces (Semester 2) — 74 /
47 / 30. Other items: Head of Subject (40% FTE) 634, General Administration 79,
RSE 491.04, Preparation time 90, Off-campus (Other) delivery 30, PGR Supervision
(2nd supervisor) 20. These sum exactly to the category totals above.
