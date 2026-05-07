#!/usr/bin/env python3
"""
step3_extract_to_xlsx.py
========================
Step 3 of the LAUSD procurement pipeline. Reads every PDF in `procurement_pages/`,
extracts each contract item, and appends new rows to a master Excel file.

PDFs that already appear in the spreadsheet (matched by 'Source PDF' column)
are skipped, so this script is safe to run repeatedly — only NEW PDFs add rows.

Usage:
    python step3_extract_to_xlsx.py
    python step3_extract_to_xlsx.py --indir procurement_pages \
        --xlsx LAUSD_Procurement_Contracts.xlsx

Requirements:
    pip install openpyxl pdfplumber  (pdfplumber pulls in pypdf, pdfminer)
    Also requires `pdftotext` from poppler-utils on the system PATH.
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

DEFAULT_INDIR = Path("procurement_pages")
DEFAULT_XLSX = Path("LAUSD_Procurement_Contracts.xlsx")

HEADERS = [
    "Board Meeting Date", "Board Report #", "Item", "Contractor", "Contract ID",
    "Division", "Contract Type", "Term Start", "Term End", "Term Notes",
    "Amount ($)", "Funding Source", "Summary", "Initial Value",
    "New Aggregate Value", "Source PDF",
]

FIELD_KEYS = [
    "board_meeting_date", "board_report_number", "item", "contractor", "contract_id",
    "division", "contract_type", "term_start", "term_end", "term_notes",
    "amount", "funding_source", "summary", "initial_value",
    "new_aggregate_value", "source_pdf",
]


# --------------- Parsing logic (same approach as backfill) -----------------

DIVISION_KEYWORDS = (
    "DIVISION OF INSTRUCTION", "DIVISION OF SCHOOL OPERATIONS",
    "DIVISION OF SPECIAL EDUCATION", "DIVISION OF DISTRICT OPERATIONS",
    "DIVISION OF ADULT AND CAREER EDUCATION", "INFORMATION TECHNOLOGY DIVISION",
    "INFORMATION TECHNOLOGY SERVICES", "FACILITIES SERVICES DIVISION",
    "FOOD SERVICES DIVISION", "TRANSPORTATION SERVICES DIVISION",
    "TRANSPORATION SERVICES DIVISION", "TRANSPORTATION SERVICES BRANCH",
    "PROCUREMENT SERVICES DIVISION", "OFFICE OF THE CHIEF MEDICAL DIRECTOR",
    "OFFICE OF THE CHIEF FINANCIAL OFFICER", "OFFICE OF THE CHIEF STRATEGY OFFICER",
    "OFFICE OF THE CHIEF FACILITIES EXECUTIVE", "OFFICE OF DATA AND ACCOUNTABILITY",
    "OFFICE OF EMERGENCY MANAGEMENT", "SCHOOL CULTURE, CLIMATE AND SAFETY",
    "SCHOOL CULTURE CLIMATE AND SAFETY", "LOS ANGELES SCHOOL POLICE DEPARTMENT",
    "MAINTENANCE & OPERATIONS", "MAINTENANCE AND OPERATIONS",
    "EARLY CHILDHOOD EDUCATION DIVISION", "MULTILINGUAL MULTICULTURAL EDUCATION DEPARTMENT",
    "STUDENT HEALTH AND HUMAN SERVICES", "BEYOND THE BELL BRANCH",
    "INTERSCHOLASTIC ATHLETIC DEPARTMENT", "HUMAN RESOURCES DIVISION",
    "OFFICE OF SCHOOL OPERATIONS", "BUSINESS SERVICES DIVISION",
)


def pdf_to_text(pdf_path: Path) -> str:
    """Run `pdftotext -layout` and return the text."""
    try:
        out = subprocess.run(
            ["pdftotext", "-layout", str(pdf_path), "-"],
            capture_output=True, text=True, timeout=60, check=True,
        )
        return out.stdout
    except Exception as e:
        print(f"  ! pdftotext failed for {pdf_path.name}: {e}")
        return ""


def parse_date(text):
    m = re.search(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b",
        text)
    return f"{m.group(1)} {int(m.group(2))}, {m.group(3)}" if m else ""


def parse_report_no(text):
    m = re.search(r"Bd\.\s*of\s*Ed\s*Rpt\.\s*No\.?\s*([\d]+-\d+/\d+)", text)
    return m.group(1) if m else ""


def normalize_date(d):
    if not d or "/" not in d:
        return d
    parts = d.split("/")
    if len(parts) == 3:
        mm, dd, yy = parts
        if len(yy) == 2:
            yr = int(yy)
            yy = f"20{yy}" if yr < 80 else f"19{yy}"
        return f"{mm.zfill(2)}/{dd.zfill(2)}/{yy}"
    return d


def parse_term(block):
    m = re.search(
        r"Contract\s+Term:\s*(\d{2}/\d{2}/\d{2,4})\s*(?:through|–|-)\s*(\d{2}/\d{2}/\d{2,4})",
        block)
    if m:
        return normalize_date(m.group(1)), normalize_date(m.group(2))
    m = re.search(r"(\d{2}/\d{2}/\d{2,4})\s*[–\-]\s*(\d{2}/\d{2}/\d{2,4})", block)
    if m:
        return normalize_date(m.group(1)), normalize_date(m.group(2))
    if "One-time" in block or "one-time" in block:
        return "One-time", "One-time"
    return "", ""


def parse_term_notes(block):
    notes = []
    if re.search(r"includes?\s+(?:a\s+)?(?:\w+[-\s])?(\d+|one|two|three|four|five)[-\s]?\(?\d*\)?[-\s]?year[\s\w]*(?:renewal|extension)",
                 block, re.I):
        m = re.search(r"includes?\s+([\w\(\)\s\-]+?(?:renewal|extension)[\w\s]*?)(?:\.|\n|$)",
                      block, re.I)
        if m:
            notes.append("Includes " + re.sub(r"\s+", " ", m.group(1).strip()))
    if re.search(r"capacity\s+increase|increase\s+(?:the\s+)?capacity", block, re.I):
        notes.append("Capacity increase")
    if re.search(r"piggyback|piggy\s*back", block, re.I) and "piggyback" not in " ".join(notes).lower():
        notes.append("Piggyback contract")
    return "; ".join(dict.fromkeys(notes))


def parse_amount(block):
    matches = re.findall(r"\$([\d,]+(?:\.\d+)?)", block)
    amounts = []
    for m in matches:
        val = m.replace(",", "")
        try:
            n = int(float(val))
            if n >= 1000:
                amounts.append(n)
        except ValueError:
            continue
    return max(amounts) if amounts else 0


def parse_initial_aggregate(block):
    init = agg = None
    m = re.search(r"Initial\s+(?:Authorized|Contract)\s+Value:?\s*\$?([\d,]+)", block)
    if m:
        try:
            init = int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    m = re.search(r"Aggregate(?:[\w\s\(\)]+?)Value(?:\s*For[\w\s\(\)]+)?:\s*\$?([\d,]+)", block)
    if m:
        try:
            agg = int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    m = re.search(r"New\s+Aggregate\s+Contract\s+Value:?\s*\$?([\d,]+)", block)
    if m:
        try:
            agg = int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    return init, agg


def parse_funding(block):
    candidates = []
    for m in re.finditer(
        r"([A-Z][\w\s,/&\-]+?(?:Funds|Fund|Grant|Revenue|Funding|RSI|ELOP|ESSER\s*III|COVID-?19|COPS))\s*\(\s*(\d{1,3})\s*%\s*\)",
        block):
        name = re.sub(r"\s+", " ", m.group(1).strip())
        if len(name) <= 80:
            candidates.append(f"{name} ({m.group(2)}%)")
    if candidates:
        return "; ".join(dict.fromkeys(candidates))
    if "Not applicable" in block or "not applicable" in block.lower():
        return "Not applicable"
    if "Revenue" in block and re.search(r"<\$[\d,]+>", block):
        return "Revenue"
    return ""


def find_division_in_block(text):
    for line in text.split("\n"):
        s = re.sub(r"\$[\d,]+", "", line.strip()).strip()
        s = re.sub(r"\(Cont\.\)", "", s).strip()
        for kw in DIVISION_KEYWORDS:
            if s.startswith(kw):
                t = kw.title().replace(" Of ", " of ").replace(" And ", " and ").replace(" The ", " the ")
                if "(Cont.)" in line:
                    t += " (Cont.)"
                return t
    return ""


def parse_contract_type(block):
    section_text = block[:2000].upper()
    base = "Professional Services"
    for k, v in (
        ("PROFESSIONAL SERVICE", "Professional Services"),
        ("GOODS AND GENERAL SERVICES", "Goods/General Services"),
        ("DONATION", "Donation of Surplus Equipment"),
        ("REVENUE", "Professional Services"),
    ):
        if k in section_text:
            base = v
            break
    block_l = block.lower()
    suffix = ""
    if "single-source" in block_l or "sole-source" in block_l or "sole source" in block_l:
        suffix = " (Single-Source)"
    elif "amendment" in block_l or "increase capacity" in block_l or "increase the capacity" in block_l:
        suffix = " (Amendment)"
    elif "extension" in block_l and "extension agreement" not in block_l and "year extension" not in block_l:
        suffix = " (Extension)"
    elif "master agreement" in block_l:
        suffix = " (Master Agreement)"
    return base + suffix


def first_paragraph_summary(block):
    m = re.search(r"((?:Approval|Authorization|Approval to|Authorization to)[^\n]*)", block)
    if not m:
        return ""
    rest = block[m.start():m.start() + 500]
    sent = re.split(r"(?<=[.!?])\s+", rest, maxsplit=1)
    return re.sub(r"\s+", " ", sent[0].strip())[:300]


def extract_contractor_id_legacy(block):
    header_match = re.search(r"CONTRACTOR\s+IDENTIFI", block, re.I)
    if not header_match:
        header_match = re.search(r"^\s*CONTRACTOR\s", block, re.M | re.I)
    if not header_match:
        return None, None
    after = block[header_match.start():]
    desc_match = re.search(r"\n\s*(?:Approval|Authorization|Approval to|Authorization to)", after)
    table_text = after[:desc_match.start()] if desc_match else after[:1500]
    contractor_chunks, id_chunks = [], []
    for line in table_text.split("\n")[1:]:
        ls = line.strip()
        if not ls or any(x in ls.upper() for x in
                         ("CONTRACTOR", "IDENTIFI", "FUNDS", "AMOUNT", "DESCRIPTION",
                          "CATION", "ATTACHMENT", "REQUEST FOR", "DELEGATED", "Bd. of Ed",
                          "PAGE", "BOARD OF EDUCATION", "ITEM ", "EXCEEDING", "CAPACITY")):
            continue
        parts = re.split(r"\s{3,}", line.rstrip())
        parts = [p.strip() for p in parts if p.strip()]
        if not parts:
            continue
        first = parts[0]
        if first.isupper() and len(first) > 5 and not any(t in first for t in ("INC", "LLC", "CORP", "DBA")):
            continue
        if re.match(r"^[\$\(\d\.,\)\%\s]+$", first):
            continue
        if first.lower() in ("general", "funds", "esser", "bond", "revenue", "grant", "covid-19", "categorical"):
            if "Various" in first:
                contractor_chunks.append(first)
            continue
        contractor_chunks.append(first)
        if len(parts) > 1:
            second = parts[1]
            if re.match(r"^(?:4400|45007|\d{4}-\d|\(IFB|\(RFP|C\d|CA DGS)", second) or re.match(r"^\d{6,}$", second):
                id_chunks.append(second)
    contractor = re.sub(r"\s+", " ", " ".join(contractor_chunks)).strip()[:500]
    contract_id = "; ".join(dict.fromkeys(id_chunks))[:300]
    return contractor or None, contract_id or None


def extract_contractor_id_new(block):
    contractor = contract_id = None
    for m in re.finditer(r"([A-Z][\w\s,\&\.\(\)\-/'’]+?)\s*/\s*(C\d+(?:[\s\-,;]+C?\d+)*)", block):
        contractor = m.group(1).strip()
        contract_id = m.group(2).strip()
        break
    if "Various Vendors" in block:
        m = re.search(r"\*\s*(.+?)(?:\n\s*\n|\n\s*\*\s*\n|\n\s*Item|\n\s*•|\n\s*\.\s|$)",
                      block, re.S)
        if m:
            footnote = re.sub(r"\s+", " ", m.group(1).strip())
            if len(footnote) > 30:
                contractor = footnote[:2000]
    return contractor, contract_id


def clean_contractor(c):
    if not c:
        return ""
    c = re.sub(r"\b4400\d{6,}\b", "", c)
    c = re.sub(r"\b4500\d{6,}\b", "", c)
    c = re.sub(r"\bC\d{4,5}\b", "", c)
    c = re.sub(r"\(\d{1,3}%\)", "", c)
    c = re.sub(r"\b(?:General Funds|Bond Funds|ESSER III|ESSER|Revenue|Funds|Or Office|Office)\b\s*$", "", c, flags=re.I)
    c = re.sub(r"\bvarious\s+per\s+requesting.*", "", c, flags=re.I)
    c = re.sub(r"\bELOP\b", "", c)
    c = re.sub(r"\$[\d,]+", "", c)
    return re.sub(r"\s+", " ", c).strip().rstrip(",;: -")


def find_item_blocks(text):
    item_positions = []
    for m in re.finditer(r"(?:^|\n)\s*Item\s+([A-Z])\s*\n", text):
        item_positions.append((m.start(), m.group(1)))
    if not item_positions:
        for m in re.finditer(r"\bItem\s+([A-Z])\b", text):
            item_positions.append((m.start(), m.group(1)))
    seen, unique = set(), []
    for pos, letter in item_positions:
        if letter not in seen:
            seen.add(letter)
            unique.append((pos, letter))
    unique.sort()
    blocks = []
    for i, (pos, letter) in enumerate(unique):
        end = unique[i + 1][0] if i + 1 < len(unique) else len(text)
        block_start = max(0, pos - 3500)
        section_starts = [m.start() for m in re.finditer(r"\n\s*[A-D]\.\s+APPROVAL", text[:pos])]
        if section_starts:
            block_start = section_starts[-1]
        if i > 0 and unique[i - 1][0] > block_start:
            block_start = unique[i - 1][0] + 1
        blocks.append((letter, text[block_start:end]))
    return blocks


def extract_pdf(pdf_path: Path) -> list[dict]:
    text = pdf_to_text(pdf_path)
    if not text:
        return []
    rpt = parse_report_no(text)
    date = parse_date(text)
    rows = []
    for letter, block in find_item_blocks(text):
        c, cid = extract_contractor_id_new(block)
        if not c:
            c2, cid2 = extract_contractor_id_legacy(block)
            if c2:
                c, cid = c2, cid2
        ts, te = parse_term(block)
        init, agg = parse_initial_aggregate(block)
        rows.append({
            "board_meeting_date": date,
            "board_report_number": rpt,
            "item": letter,
            "contractor": clean_contractor(c or ""),
            "contract_id": cid or "",
            "division": find_division_in_block(block),
            "contract_type": parse_contract_type(block),
            "term_start": ts,
            "term_end": te,
            "term_notes": parse_term_notes(block),
            "amount": parse_amount(block),
            "funding_source": parse_funding(block),
            "summary": first_paragraph_summary(block),
            "initial_value": init,
            "new_aggregate_value": agg,
            "source_pdf": pdf_path.name,
        })
    return rows


# --------------- xlsx append --------------------------------------------------

def load_existing_pdfs(xlsx_path: Path) -> set[str]:
    """Read 'Source PDF' column from the workbook to know what's already in there."""
    if not xlsx_path.exists():
        return set()
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if "Contracts" not in wb.sheetnames:
        return set()
    ws = wb["Contracts"]
    pdfs = set()
    src_col = HEADERS.index("Source PDF")
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue
        if row and len(row) > src_col and row[src_col]:
            pdfs.add(str(row[src_col]).strip())
    return pdfs


def append_rows(xlsx_path: Path, new_rows: list[dict]):
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
        if "Contracts" not in wb.sheetnames:
            ws = wb.create_sheet("Contracts", 0)
            _write_headers(ws)
        else:
            ws = wb["Contracts"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Contracts"
        _write_headers(ws)

    body_font = Font(name="Arial", size=10)
    body_align = Alignment(vertical="top", wrap_text=True)
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    next_row = ws.max_row + 1
    for r_idx, row_dict in enumerate(new_rows, next_row):
        for c_idx, key in enumerate(FIELD_KEYS, 1):
            val = row_dict.get(key)
            if val == "":
                val = None
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = body_font
            cell.alignment = body_align
            cell.border = border
        for col in (11, 14, 15):
            ws.cell(row=r_idx, column=col).number_format = '"$"#,##0;[Red]("$"#,##0);"-"'
        contractor_len = len(str(row_dict.get("contractor") or ""))
        if contractor_len > 800:
            ws.row_dimensions[r_idx].height = 220
        elif contractor_len > 300:
            ws.row_dimensions[r_idx].height = 110
        elif contractor_len > 150:
            ws.row_dimensions[r_idx].height = 75
        else:
            ws.row_dimensions[r_idx].height = 60

    ws.freeze_panes = "D2"
    wb.save(xlsx_path)


def _write_headers(ws):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F4E78")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_align
        c.border = border
    widths = {"A": 17, "B": 12, "C": 6, "D": 38, "E": 22, "F": 22, "G": 22,
              "H": 12, "I": 12, "J": 28, "K": 14, "L": 28, "M": 50,
              "N": 14, "O": 16, "P": 38}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 36


# --------------- main --------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Append new procurement PDFs to the master Excel.")
    parser.add_argument("--indir", default=str(DEFAULT_INDIR),
                        help=f"Folder of procurement PDFs (default: {DEFAULT_INDIR})")
    parser.add_argument("--xlsx", default=str(DEFAULT_XLSX),
                        help=f"Master spreadsheet path (default: {DEFAULT_XLSX})")
    args = parser.parse_args()

    indir = Path(args.indir)
    xlsx_path = Path(args.xlsx)

    if not indir.exists():
        print(f"Input folder '{indir}' not found.")
        sys.exit(1)

    pdfs = sorted(indir.glob("*.pdf"))
    if not pdfs:
        print(f"No PDFs in '{indir}'.")
        sys.exit(0)

    already = load_existing_pdfs(xlsx_path)
    print(f"Master xlsx: {xlsx_path} ({'exists' if xlsx_path.exists() else 'will be created'})")
    print(f"Already-processed PDFs: {len(already)}")

    new_pdfs = [p for p in pdfs if p.name not in already]
    print(f"New PDFs to process: {len(new_pdfs)}")

    all_new_rows = []
    for pdf in new_pdfs:
        rows = extract_pdf(pdf)
        if rows:
            print(f"  + {pdf.name}: {len(rows)} contract item(s)")
            all_new_rows.extend(rows)
        else:
            print(f"  - {pdf.name}: no items extracted")

    if all_new_rows:
        append_rows(xlsx_path, all_new_rows)
        print(f"\nAppended {len(all_new_rows)} new row(s) to '{xlsx_path}'.")
    else:
        print("\nNo new rows to add.")


if __name__ == "__main__":
    main()
