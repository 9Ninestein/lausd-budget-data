#!/usr/bin/env python3
"""
Extract structured data from procurement contract PDFs and write to Excel.

Reads every PDF in the procurement_pages/ folder, extracts one row per
contract entry (contractor, contract IDs, amount, term, source of funds,
requester, division, meeting date, report number) and writes to a .xlsx file.

Usage:
    python extract_procurement.py
    python extract_procurement.py --indir procurement_pages --out contracts.xlsx

Requirements:
    pip install pymupdf openpyxl
"""

import argparse
import re
import sys
from pathlib import Path

import fitz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DEFAULT_INDIR = Path("procurement_pages")
DEFAULT_OUT   = Path("contracts.xlsx")

# ── regex helpers ────────────────────────────────────────────────────────────

DATE_PAT      = re.compile(
    r'\b(January|February|March|April|May|June|July|August'
    r'|September|October|November|December)\s+\d{1,2},\s+\d{4}', re.I)
RPT_PAT       = re.compile(r'Bd\.?\s*of\s*Ed\.?\s*Rpt\.?\s*No\.?\s*([^\s\n]+)', re.I)
# Old format: 4400012153  |  New format (2025+): C6751
CONTRACT_ID   = re.compile(r'\b(44\d{8}|C\d{3,6})\b')
RFP_ID        = re.compile(r'\(RFP\s+\d+\)')
AMOUNT_PAT    = re.compile(r'\$([\d,]+(?:\.\d+)?)')
TERM_PAT      = re.compile(
    r'Contract\s+Term[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})\s+through\s+(\d{1,2}/\d{1,2}/\d{2,4})', re.I)
# New-format term is "MM/DD/YY – MM/DD/YY" (en-dash, in its own column)
TERM_COL_PAT  = re.compile(r'(\d{1,2}/\d{1,2}/\d{2,4})\s*[–\-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})')
REQUESTER_PAT = re.compile(r'Requester[s]?[:\s]+([^\n]+)', re.I)
SECTION_PAT   = re.compile(r'\b([A-Z])\.\s+(APPROVAL OF[^\n]+)', re.M)
DIVISION_PAT  = re.compile(
    r'(?:DIVISION OF [A-Z ,/&()\-]+|[A-Z ,/&()\-]+ SERVICES?'
    r'|[A-Z ,/&()\-]+ DEPARTMENT)', re.M)

# Column header text to filter out of extracted fields
HEADER_NOISE = re.compile(
    r'^(CONTRACTOR|IDENTIFI|CATION|NO\.|SOURCE|OF|FUNDS|AMOUNT'
    r'|DESCRIPTION|ATTACHMENT|REQUEST FOR|DELEGATED|PAGE\s*\d|BD\.|BOARD)', re.I)


# ── per-page extraction ──────────────────────────────────────────────────────

def extract_page(page, source_file: str) -> list[dict]:
    rect       = page.rect
    page_width = rect.width
    full_text  = page.get_text()
    blocks     = page.get_text("blocks")   # (x0,y0,x1,y1,text,block_no,block_type)

    # ── page-level metadata ──────────────────────────────────────────────────
    date_m  = DATE_PAT.search(full_text)
    meeting_date = date_m.group(0) if date_m else ""

    rpt_m   = RPT_PAT.search(full_text)
    report  = rpt_m.group(1).rstrip(",") if rpt_m else ""

    sect_m  = SECTION_PAT.search(full_text)
    section = f"{sect_m.group(1)}. {sect_m.group(2).strip()}" if sect_m else ""

    # Division: find the capitalised line that names the organisational unit
    div_candidates = DIVISION_PAT.findall(full_text)
    division = div_candidates[0].strip() if div_candidates else ""

    # ── build positioned block list ──────────────────────────────────────────
    positioned = []
    for b in blocks:
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        text = text.strip()
        if text and not HEADER_NOISE.match(text):
            positioned.append((x0, y0, x1, y1, text))

    RIGHT_MIN = page_width * 0.72

    # ── find the y-position of the column-header row ─────────────────────────
    # Everything ABOVE the header row is page metadata; amounts there are
    # division totals, not individual contract amounts — we skip them.
    header_y = 0
    for x0, y0, x1, y1, text in positioned:
        if re.search(r'\bCONTRACTOR\b', text, re.I) and y0 > header_y:
            header_y = y0

    # ── use right-column amount blocks as row anchors ─────────────────────────
    # Every format has exactly one amount block per contract in the right column.
    amount_anchors = []
    for x0, y0, x1, y1, text in positioned:
        if x0 >= RIGHT_MIN and y0 > header_y:
            amt_m = AMOUNT_PAT.search(text)
            if amt_m:
                amount_anchors.append((y0, "$" + amt_m.group(1)))

    if not amount_anchors:
        return []

    # ── collect footnote contract IDs (sometimes listed at page bottom) ───────
    # Pattern: "**CONTRACT NOS: C9552; C9880 through C9893 ..."
    # Also just any full-width block below the table body containing C####
    footnote_ids: list[str] = []
    FOOTNOTE_NOS = re.compile(r'CONTRACT\s+NOS?[:\s]+(.+)', re.I | re.S)
    for x0, y0, x1, y1, text in positioned:
        fn_m = FOOTNOTE_NOS.search(text)
        if fn_m:
            raw = fn_m.group(1)
            # expand "C9880 through C9893" into individual IDs
            for rng in re.finditer(r'(C\d{3,6})\s+through\s+(C\d{3,6})', raw, re.I):
                start = int(rng.group(1)[1:])
                end   = int(rng.group(2)[1:])
                footnote_ids.extend(f"C{n}" for n in range(start, end + 1))
            # individual IDs not part of a range
            in_ranges = {m.group(0) for m in re.finditer(
                r'C\d{3,6}\s+through\s+C\d{3,6}', raw, re.I)}
            for single in CONTRACT_ID.findall(raw):
                context = re.search(rf'{re.escape(single)}\s+through|through\s+{re.escape(single)}', raw, re.I)
                if not context:
                    footnote_ids.append(single)

    # ── for each amount anchor, find all other fields ─────────────────────────
    contract_entries = []
    for anchor_y, amount in amount_anchors:
        Y_BAND = 100  # vertical tolerance in points

        # All non-right blocks within the y-band
        row_blocks = [
            (x0, y0, text) for x0, y0, x1, y1, text in positioned
            if abs(y0 - anchor_y) < Y_BAND and x0 < RIGHT_MIN
        ]
        if not row_blocks:
            continue

        # Sort by x so leftmost = contractor, rightmost = source of funds
        row_blocks.sort(key=lambda t: t[0])

        # Leftmost block → contractor name + possible contract IDs
        left_text = row_blocks[0][2] if row_blocks else ""
        ids = CONTRACT_ID.findall(left_text)
        pre = CONTRACT_ID.split(left_text)[0]
        pre = RFP_ID.sub("", pre).strip().strip("*").strip()
        pre = re.sub(r'\bItem\s+[A-Z]\b', "", pre, flags=re.I).strip()
        contractor = pre if pre else "Various Vendors"

        # If no IDs in left block, check a dedicated ID column (second-from-left)
        if not ids and len(row_blocks) > 1:
            ids = CONTRACT_ID.findall(row_blocks[1][2])

        # Still no IDs → fall back to footnote IDs (Various Vendors bench)
        if not ids and footnote_ids:
            ids = footnote_ids  # shared pool for this page

        # Contract term: look for date–date pattern in any middle block
        contract_term = ""
        for x0, y0, text in row_blocks[1:]:
            tm = TERM_COL_PAT.search(text)
            if tm:
                contract_term = f"{tm.group(1)} through {tm.group(2)}"
                break
        # Fall back to in-description term (old format)
        if not contract_term:
            tm = TERM_PAT.search(full_text)
            if tm:
                contract_term = f"{tm.group(1)} through {tm.group(2)}"

        # Source of funds: rightmost block before the amount column
        source_of_funds = ""
        if len(row_blocks) > 1:
            # exclude left (contractor) block; pick the rightmost remaining
            candidates = [t for x0, y0, t in row_blocks[1:]
                          if not CONTRACT_ID.search(t) and not TERM_COL_PAT.search(t)]
            if candidates:
                source_of_funds = re.sub(r'\s+', " ", candidates[-1]).strip()

        contract_entries.append({
            "source_file":     source_file,
            "meeting_date":    meeting_date,
            "report_number":   report,
            "section":         section,
            "division":        division,
            "contractor":      contractor,
            "contract_ids":    ", ".join(ids),
            "amount":          amount,
            "source_of_funds": source_of_funds,
            "contract_term":   contract_term,
            "requester":       "",
        })

    # ── assign requesters in order of appearance ─────────────────────────────
    requesters = [m.strip() for m in REQUESTER_PAT.findall(full_text)]
    for i, entry in enumerate(contract_entries):
        if i < len(requesters):
            entry["requester"] = requesters[i]

    return contract_entries


# ── Excel writer ─────────────────────────────────────────────────────────────

COLUMNS = [
    ("Source File",       25),
    ("Meeting Date",      18),
    ("Report Number",     16),
    ("Section",           35),
    ("Division",          35),
    ("Contractor",        35),
    ("Contract ID(s)",    30),
    ("Amount",            14),
    ("Source of Funds",   30),
    ("Contract Term",     28),
    ("Requester",         35),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def write_xlsx(rows: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Procurement Contracts"

    # Header row
    for col_idx, (header, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    field_keys = [
        "source_file", "meeting_date", "report_number", "section",
        "division", "contractor", "contract_ids", "amount",
        "source_of_funds", "contract_term", "requester",
    ]

    for row_idx, row in enumerate(rows, 2):
        for col_idx, key in enumerate(field_keys, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(key, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--indir", default=str(DEFAULT_INDIR))
    parser.add_argument("--out",   default=str(DEFAULT_OUT))
    args   = parser.parse_args()

    indir    = Path(args.indir)
    out_path = Path(args.out)

    if not indir.exists():
        print(f"Folder '{indir}' not found.")
        sys.exit(1)

    pdfs = sorted(indir.glob("*.pdf"))
    if not pdfs:
        print(f"No PDFs found in '{indir}'.")
        sys.exit(1)

    print(f"Extracting data from {len(pdfs)} PDF(s)...\n")

    all_rows = []
    for pdf_path in pdfs:
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"  ERROR opening {pdf_path.name}: {e}")
            continue

        page_rows = []
        for page in doc:
            page_rows.extend(extract_page(page, pdf_path.name))
        doc.close()

        print(f"  {pdf_path.name}: {len(page_rows)} contract(s)")
        all_rows.extend(page_rows)

    if not all_rows:
        print("\nNo contract data extracted.")
        sys.exit(1)

    write_xlsx(all_rows, out_path)
    print(f"\nDone. {len(all_rows)} rows written to {out_path}")


if __name__ == "__main__":
    main()
