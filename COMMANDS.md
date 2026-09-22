# Common commands

Quick reference for this directory. Run everything from here, with `python3`
(not `python`). `WORKBOOK` = the current `.xlsm` master, e.g.:

```bash
WORKBOOK="2026-27 Resources Planner - School of Computer Science (updated 2026-07-1) (1).xlsm"
```

If figures come back 0/blank, the workbook's formula cache is stale — open and
re-save in Excel, or force a recalc:

```bash
soffice --headless --convert-to xlsx --calc "$WORKBOOK"
```

## Per-staff allocation PDFs

```bash
# whole department
python3 generate_allocation_pdfs.py "$WORKBOOK" -o allocation_pdfs --skip-empty

# single person, fast iteration
python3 generate_allocation_pdfs.py "$WORKBOOK" --id NHQR5 -o preview
```

## Per-module PDFs

```bash
# all modules
python3 generate_allocation_pdfs.py "$WORKBOOK" -o module_pdfs --modules

# single module
python3 generate_allocation_pdfs.py "$WORKBOOK" --module KV6013 -o module_pdfs --modules
```

## Admin sheet / doc

```bash
python3 generate_admin_sheet.py "$WORKBOOK" -o admin_sheet.xlsx
python3 generate_admin_doc.py "$WORKBOOK" -o admin_doc.docx
```

## Emailing PDFs (macOS Mail.app)

```bash
# staff allocation PDFs — always dry-run first
python3 send_allocation_emails.py allocation_pdfs --dry-run
python3 send_allocation_emails.py allocation_pdfs

# module PDFs to tutors
python3 email_module_pdfs.py "$WORKBOOK" --dry-run
python3 email_module_pdfs.py "$WORKBOOK"

# just one person / module
python3 send_allocation_emails.py allocation_pdfs --id NHQR5
python3 email_module_pdfs.py "$WORKBOOK" --module KV6013
```

See [README.md](README.md) for the full spreadsheet layout and script internals.
