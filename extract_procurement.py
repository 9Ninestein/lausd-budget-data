#!/usr/bin/env python3
"""
Extract structured data from procurement contract PDFs and write to Excel.

Columns: Source File, Meeting Date, Report Number, Division, Contractor(s),
         Contract ID(s), Contract Term, Description, Source of Funds,
         Amount, Aggregate Value

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
    r'|September|October|November|December)\s+\d{1,2},?\s+\d{4}', re.I)
RPT_PAT       = re.compile(r'Bd\.?\s*of\s*Ed\.?\s*Rpt\.?\s*No\.?\s*([^\s\n]+)', re.I)
CONTRACT_ID   = re.compile(r'\b(44\d{8}|C\d{3,6})\b')
RFP_ID        = re.compile(r'\(?\bRFP\s+\d+\)?', re.I)
AMOUNT_PAT    = re.compile(r'\$([\d,]+(?:\.\d+)?)')
AGGREGATE_PAT = re.compile(r'Aggregate[^\n$]*\$([\d,]+(?:\.\d+)?)', re.I)
TERM_PAT      = re.compile(
    r'Contract\s+Term[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})\s+through\s+(\d{1,2}/\d{1,2}/\d{2,4})', re.I)
TERM_COL_PAT  = re.compile(r'(\d{1,2}/\d{1,2}/\d{2,4})\s*[–\-]\s*(\d{1,2}/\d{1,2}/\d{2,4})')
REQUESTER_PAT = re.compile(r'Requester[s]?[:\s]+([^\n]+)', re.I)

# Bold text we always ignore (fixed page headers)
FIXED_HEADERS = re.compile(
    r'^(ATTACHMENT|REQUEST FOR APPROVAL|DELEGATED AUTHORITY'
    r'|NEW CONTRACTS|APPROVAL OF PROFESSIONAL|AUTHORIZATION TO INCREASE'
    r'|EXCEEDING|\$\d)', re.I)

# Division / org-unit keywords
DIVISION_KW   = re.compile(
    r'\b(DIVISION|OFFICE|SERVICES|SERVICE|BRANCH|DEPARTMENT'
    r'|COMMISSION|CENTER|CENTRE|BUREAU|PROGRAM)\b', re.I)

# Column-header noise to exclude from description / source
HEADER_NOISE  = re.compile(
    r'^(CONTRACTOR|IDENTIFI|CATION|NO\.|SOURCE|OF|FUNDS|AMOUNT'
    r'|DESCRIPTION|CONTRACT\s+TERM|ATTACHMENT|REQUEST FOR'
    r'|DELEGATED|PAGE\s*\d|BD\.|BOARD)', re.I)


# ── helpers ──────────────────────────────────────────────────────────────────

def is_bold(span: dict) -> bool:
    return bool(span['flags'] & 16) or 'Bold' in span.get('font', '')


def bold_lines(page) -> list[tuple[float, float, str]]:
    """Return (x0, y0, text) for every line that is predominantly bold."""
    results = []
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            spans = line.get('spans', [])
            if not spans:
                continue
            line_text = ''.join(s['text'] for s in spans).strip()
            bold_chars = sum(len(s['text']) for s in spans if is_bold(s))
            if line_text and bold_chars / max(len(line_text), 1) > 0.5:
                x0 = line['bbox'][0]
                y0 = line['bbox'][1]
                results.append((x0, y0, line_text))
    return results


def expand_various_vendors(page, page_width: float) -> list[str]:
    """
    When contractor is 'Various Vendors', find the asterisk footnote that
    lists individual vendor names and return them as a list.
    """
    vendors: list[str] = []
    blocks = page.get_text('blocks')
    for b in blocks:
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        text = text.strip()
        width_frac = (x1 - x0) / page_width
        # Footnote: wide block starting with one or more asterisks
        if width_frac > 0.70 and re.match(r'^\*', text):
            # Strip leading asterisks and optional "CONTRACTORS:" label
            raw = re.sub(r'^\*+\s*(?:CONTRACTORS?[:\s]+)?', '', text, flags=re.I)
            # Remove any trailing "CONTRACT NOS:" section
            raw = re.split(r'\*+\s*CONTRACT NOS?:', raw, flags=re.I)[0]
            # Also strip second-asterisk footnotes (**existing vendors…)
            raw = re.split(r'\*\*', raw)[0]
            parts = [v.strip() for v in re.split(r';', raw) if v.strip()]
            vendors.extend(parts)
    return vendors


def extract_divisions(page, bold_text_lines: list) -> list[tuple[float, str]]:
    """Return (y0, division_name) for each division header on the page."""
    results = []
    for x0, y0, text in bold_text_lines:
        if FIXED_HEADERS.match(text):
            continue
        if DIVISION_KW.search(text):
            # Clean up: strip embedded dollar amounts
            clean = re.sub(r'\$[\d,]+', '', text).strip().strip('/')
            clean = re.sub(r'\s+', ' ', clean).strip()
            if clean:
                results.append((y0, clean))
    return results


def collect_description_blocks(page, header_y: float, page_width: float) -> list[tuple[float, str]]:
    """
    Return (y0, text) for blocks that look like description / narrative text:
    - Wide blocks (>55% of page width) below the column-header row
    - Middle-column blocks (old format) — between left and right zones
    - Exclude footnotes (start with *), page numbers, and header noise
    """
    RIGHT_MIN = page_width * 0.72
    results = []
    for b in page.get_text('blocks'):
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        text = text.strip()
        if not text or y0 <= header_y:
            continue
        if re.match(r'^\*', text):       # footnote
            continue
        if re.match(r'^\s*\d+\s*$', text):  # page number
            continue
        if HEADER_NOISE.match(text):
            continue
        if x0 >= RIGHT_MIN:              # amount column
            continue
        width_frac = (x1 - x0) / page_width
        # Wide paragraph (new format) OR middle-column description (old format)
        if width_frac > 0.55 or (0.25 < x0 / page_width < 0.65 and width_frac > 0.25):
            results.append((y0, text))
    # Remove duplicates while preserving order
    seen = set()
    deduped = []
    for y0, text in results:
        if text not in seen:
            seen.add(text)
            deduped.append((y0, text))
    return sorted(deduped, key=lambda t: t[0])


# ── per-page extraction ──────────────────────────────────────────────────────

def extract_page(page, source_file: str) -> list[dict]:
    rect       = page.rect
    page_width = rect.width
    full_text  = page.get_text()
    blocks_raw = page.get_text('blocks')

    # ── page-level metadata ──────────────────────────────────────────────────
    date_m  = DATE_PAT.search(full_text)
    meeting_date = date_m.group(0) if date_m else ""

    rpt_m   = RPT_PAT.search(full_text)
    report  = rpt_m.group(1).rstrip(",") if rpt_m else ""

    # ── build positioned block list ──────────────────────────────────────────
    positioned = []
    for b in blocks_raw:
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        text = text.strip()
        if text and not HEADER_NOISE.match(text):
            positioned.append((x0, y0, x1, y1, text))

    RIGHT_MIN = page_width * 0.72

    # ── find y-position of column-header row (skip division totals above it) ─
    header_y = 0
    for x0, y0, x1, y1, text in positioned:
        if re.search(r'\bCONTRACTOR\b', text, re.I) and y0 > header_y:
            header_y = y0

    # ── bold lines for division detection ────────────────────────────────────
    b_lines   = bold_lines(page)
    divisions = extract_divisions(page, b_lines)  # [(y0, name), ...]

    # ── footnote vendor lists ─────────────────────────────────────────────────
    footnote_vendors = expand_various_vendors(page, page_width)

    # ── footnote contract IDs (various-vendor bench in newest format) ─────────
    footnote_ids: list[str] = []
    FOOTNOTE_NOS = re.compile(r'CONTRACT\s+NOS?[:\s]+(.+)', re.I | re.S)
    for x0, y0, x1, y1, text in positioned:
        fn_m = FOOTNOTE_NOS.search(text)
        if fn_m:
            raw = fn_m.group(1)
            for rng in re.finditer(r'(C\d{3,6})\s+through\s+(C\d{3,6})', raw, re.I):
                start = int(rng.group(1)[1:])
                end   = int(rng.group(2)[1:])
                footnote_ids.extend(f"C{n}" for n in range(start, end + 1))
            for single in CONTRACT_ID.findall(raw):
                if not re.search(rf'{re.escape(single)}\s+through|through\s+{re.escape(single)}', raw, re.I):
                    footnote_ids.append(single)

    # ── description blocks (shared across all contracts on page) ─────────────
    desc_blocks = collect_description_blocks(page, header_y, page_width)

    # ── use right-column amount blocks as row anchors ─────────────────────────
    amount_anchors = []
    for x0, y0, x1, y1, text in positioned:
        if x0 >= RIGHT_MIN and y0 > header_y:
            amt_m = AMOUNT_PAT.search(text)
            if amt_m:
                amount_anchors.append((y0, "$" + amt_m.group(1)))

    if not amount_anchors:
        return []

    # ── build one entry per amount anchor ────────────────────────────────────
    # Source of funds is always in a narrow right-centre zone
    SOURCE_MIN = page_width * 0.55
    contract_entries = []

    for anchor_y, amount in amount_anchors:
        Y_BAND = 100

        # Exclude very-wide blocks (>65 % of page) — these are section headers
        # or division totals, not individual column cells.
        row_blocks = sorted(
            [(x0, y0, x1, text) for x0, y0, x1, y1, text in positioned
             if abs(y0 - anchor_y) < Y_BAND
             and x0 < RIGHT_MIN
             and (x1 - x0) / page_width < 0.65],
            key=lambda t: t[0]
        )
        if not row_blocks:
            continue

        # ── contractor ───────────────────────────────────────────────────────
        left_text  = row_blocks[0][3]
        ids        = CONTRACT_ID.findall(left_text)
        pre        = CONTRACT_ID.split(left_text)[0]
        pre        = RFP_ID.sub("", pre).strip().strip("*").strip()
        pre        = re.sub(r'\bItem\s+[A-Z]\b', "", pre, flags=re.I).strip()
        contractor = pre if pre else "Various Vendors"

        # Expand "Various Vendors" using asterisk footnote
        is_various = re.search(r'various', contractor, re.I) or re.search(r'^\*+$', contractor)
        if is_various and footnote_vendors:
            contractor = "; ".join(footnote_vendors)

        # IDs: check second-from-left column if not found in left block
        if not ids and len(row_blocks) > 1:
            ids = CONTRACT_ID.findall(row_blocks[1][3])
        if not ids and footnote_ids:
            ids = footnote_ids

        # ── division: use nearest division header ABOVE this row ─────────────
        division = ""
        best_dy  = 9999
        for dy0, dname in divisions:
            diff = anchor_y - dy0
            if 0 < diff < best_dy:
                best_dy  = diff
                division = dname

        # ── contract term ─────────────────────────────────────────────────────
        contract_term = ""
        for x0, y0, x1, text in row_blocks[1:]:
            tm = TERM_COL_PAT.search(text)
            if tm:
                contract_term = f"{tm.group(1)} through {tm.group(2)}"
                break
        if not contract_term:
            tm = TERM_PAT.search(full_text)
            if tm:
                contract_term = f"{tm.group(1)} through {tm.group(2)}"

        # ── source of funds ───────────────────────────────────────────────────
        # Source is always in the right-centre zone (between description and amount)
        source_parts = [
            text for x0, y0, x1, text in row_blocks
            if x0 >= SOURCE_MIN
            and not CONTRACT_ID.search(text)
            and not TERM_COL_PAT.search(text)
        ]
        source_of_funds = re.sub(r'\s+', " ", " ".join(source_parts)).strip()

        # ── description ───────────────────────────────────────────────────────
        # Collect all description blocks that fall within this contract's
        # vertical range: from anchor_y down to the next anchor or end of page.
        next_anchors = [ay for ay, _ in amount_anchors if ay > anchor_y]
        lower_bound  = min(next_anchors) if next_anchors else page.rect.height

        desc_parts = [
            text for y0, text in desc_blocks
            if anchor_y - 20 <= y0 < lower_bound
        ]
        # In old format the description is in the middle column at the same y
        if not desc_parts:
            desc_parts = [
                text for y0, text in desc_blocks
                if abs(y0 - anchor_y) < Y_BAND
            ]
        description = "\n\n".join(dict.fromkeys(desc_parts))  # dedup, preserve order

        # ── aggregate value ───────────────────────────────────────────────────
        agg_m = AGGREGATE_PAT.search(description + "\n" + full_text)
        aggregate = f"${agg_m.group(1)}" if agg_m else ""

        contract_entries.append({
            "source_file":     source_file,
            "meeting_date":    meeting_date,
            "report_number":   report,
            "division":        division,
            "contractor":      contractor,
            "contract_ids":    ", ".join(ids),
            "contract_term":   contract_term,
            "description":     description,
            "source_of_funds": source_of_funds,
            "amount":          amount,
            "aggregate_value": aggregate,
        })

    # ── assign requesters in order of appearance ──────────────────────────────
    requesters = [m.strip() for m in REQUESTER_PAT.findall(full_text)]
    for i, entry in enumerate(contract_entries):
        if i < len(requesters):
            entry["requester"] = requesters[i]

    return contract_entries


# ── Excel writer ──────────────────────────────────────────────────────────────

COLUMNS = [
    ("Source File",       25),
    ("Meeting Date",      18),
    ("Report Number",     16),
    ("Division",          38),
    ("Contractor(s)",     45),
    ("Contract ID(s)",    28),
    ("Contract Term",     26),
    ("Description",       70),
    ("Source of Funds",   30),
    ("Amount",            14),
    ("Aggregate Value",   16),
]

FIELD_KEYS = [
    "source_file", "meeting_date", "report_number", "division",
    "contractor", "contract_ids", "contract_term", "description",
    "source_of_funds", "amount", "aggregate_value",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def write_xlsx(rows: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Procurement Contracts"

    for col_idx, (header, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    # Strip control characters that Excel rejects (keep tab, LF, CR)
    ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

    def clean(val):
        if isinstance(val, str):
            return ILLEGAL.sub('', val)
        return val

    for row_idx, row in enumerate(rows, 2):
        for col_idx, key in enumerate(FIELD_KEYS, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=clean(row.get(key, "")))
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--indir", default=str(DEFAULT_INDIR))
    parser.add_argument("--out",   default=str(DEFAULT_OUT))
    args = parser.parse_args()

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
