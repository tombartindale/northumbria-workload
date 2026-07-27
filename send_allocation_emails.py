#!/usr/bin/env python3
"""Send each staff member their workload allocation PDF via macOS Mail.app."""

import csv
import datetime
import subprocess
import sys
from pathlib import Path

EMAIL_CSV = Path(__file__).parent / "staff_emails.csv"

EMAIL_SUBJECT = "Draft 2026-27 Teaching Workload Allocation"
_TODAY = datetime.date.today().strftime("%d %B %Y")
EMAIL_BODY = f"""\
Dear {{first_name}},

Please find attached your draft workload allocation summary for 2026-27, generated on {_TODAY}.

This document is indicative, and is being shared to aid in transparency around the allocation process.

This is a brand new process, so please bear with us as we refine the document generation process.

If you have any questions about what this document shows, please don't hesitate to get in touch with myself or another subject head.

Best regards,

Tom Bartindale
Head of Subject, Computing


"""

# ── Email lookup from CSV ─────────────────────────────────────────────────────

def load_email_map() -> dict[str, str]:
    if not EMAIL_CSV.exists():
        sys.exit(f"Email map not found: {EMAIL_CSV}")
    mapping = {}
    with open(EMAIL_CSV, newline="") as f:
        for row in csv.DictReader(f):
            if row["email"].strip():
                mapping[row["staff_id"].strip()] = row["email"].strip()
    return mapping

# ── Send via Mail.app ─────────────────────────────────────────────────────────

def send_email(to_email: str, first_name: str, pdf_path: Path):
    body = EMAIL_BODY.format(first_name=first_name)
    pdf_posix = str(pdf_path.resolve())
    script = f"""
tell application "Mail"
    set msg to make new outgoing message with properties {{subject:"{EMAIL_SUBJECT}", content:"{body}", visible:false}}
    tell msg
        make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
        make new attachment with properties {{file name:POSIX file "{pdf_posix}"}} at after last paragraph
    end tell
    send msg
end tell
"""
    subprocess.run(["osascript", "-e", script], check=True, capture_output=True)

# ── Filename parsing ──────────────────────────────────────────────────────────

def staff_id_from(filename: str) -> str:
    return Path(filename).stem.split("_")[-2]

def first_name_from(filename: str) -> str:
    return Path(filename).stem.split("_")[-3]

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Email workload PDFs via macOS Mail.app")
    parser.add_argument("pdf_dir", nargs="?", default="allocation_pdfs",
                        help="Folder of PDFs (default: allocation_pdfs)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be sent without sending anything")
    parser.add_argument("--id", metavar="STAFF_ID",
                        help="Only process this one staff ID")
    args = parser.parse_args()

    pdf_dir = Path(args.pdf_dir)
    if not pdf_dir.is_dir():
        sys.exit(f"Directory not found: {pdf_dir}")

    email_map = load_email_map()

    pdfs = sorted(pdf_dir.glob("*_Allocation.pdf"))
    if args.id:
        pdfs = [p for p in pdfs if staff_id_from(p.name) == args.id.upper()]

    if not pdfs:
        sys.exit("No PDFs found.")

    print(f"Found {len(pdfs)} PDF(s).{' Dry run — no emails will be sent.' if args.dry_run else ''}\n")

    sent, skipped, failed = [], [], []

    for pdf in pdfs:
        sid   = staff_id_from(pdf.name)
        fname = first_name_from(pdf.name)
        email = email_map.get(sid)

        if not email:
            print(f"  SKIP  {sid:10s}  — no email in staff_emails.csv")
            skipped.append(sid)
            continue

        if args.dry_run:
            print(f"  DRY   {sid:10s}  {fname:15s}  → {email}")
            sent.append(sid)
            continue

        try:
            send_email(email, fname, pdf)
            print(f"  SENT  {sid:10s}  → {email}")
            sent.append(sid)
        except subprocess.CalledProcessError as exc:
            print(f"  FAIL  {sid:10s}  → {email}  [{exc.stderr.decode().strip()}]")
            failed.append(sid)

    suffix = " (dry run)" if args.dry_run else ""
    print(f"\nDone{suffix}. {len(sent)} sent, {len(skipped)} skipped (no email), {len(failed)} failed.")
    if failed:
        print("Failed IDs:", ", ".join(failed))

if __name__ == "__main__":
    main()
