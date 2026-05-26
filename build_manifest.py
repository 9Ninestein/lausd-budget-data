#!/usr/bin/env python3
"""
build_manifest.py
=================
Scan boepdfs_tabs/ and build manifest.json, extracting a title for each
tab from its first content page (the page after the "TAB N" cover).

Usage:
    python build_manifest.py
    python build_manifest.py --tabs boepdfs_tabs --out manifest.json

Requirements:
    pip install pymupdf
"""

import argparse
import json
import re
from datetime import date as _date
from pathlib import Path

# Lines that appear in the standard LAUSD Board of Education Report header
HEADER_RE = re.compile(
    r"^("
    r"Los Angeles Unified School District"
    r"|Board of Education Report"
    r"|333 South Beaudry"
    r"|Los Angeles, CA"
    r"|File #:"
    r"|Agenda Date:"
    r"|In Control:"
    r"|Version:"
    r"|powered by Legistar"
    r"|Return to Order of Business"
    r"|\d+\s*$"          # bare page numbers
    r")",
    re.I,
)

# Lines that signal the end of the title (division/action boilerplate)
STOP_RE = re.compile(
    r"^("
    r"Action Proposed|Attachment|Background|Fiscal Impact|Expected|Rationale"
    r"|Brief Description"
    r"|January|February|March|April|May|June|July|August"
    r"|September|October|November|December"
    r")",
    re.I,
)

# Short standalone division/office lines that appear after the title
TRAILING_ORG_RE = re.compile(
    r"^.{0,60}(Division|Office|Branch|Department|Disbursements)\s*$",
    re.I,
)

MONTH_NAMES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

FILENAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})_Tab_(\d+)(?:_src(\d+))?\.pdf$")


def _parse_title_from_text(text: str) -> list[str]:
    """Extract title lines from a page's text after stripping LAUSD header boilerplate."""
    parts: list[str] = []
    past_header = False
    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            if past_header and parts:
                break
            continue
        if HEADER_RE.match(line):
            past_header = True
            continue
        if not past_header:
            continue
        if STOP_RE.match(line):
            break
        parts.append(line)
        if len(" ".join(parts)) > 250:
            break
    # Trim a short trailing org-name line (e.g. "Facilities Services Division")
    if len(parts) > 1 and TRAILING_ORG_RE.match(parts[-1]):
        parts = parts[:-1]
    return parts


def extract_title(pdf_path: Path) -> str:
    import fitz

    doc = fitz.open(pdf_path)
    try:
        if len(doc) < 2:
            return "(single-page tab)"

        text = doc[1].get_text()  # page 0 = Tab cover, page 1 = first content

        # Special cases detectable from full page text
        if re.search(r"WITHDRAWN\s+PRIOR\s+TO\s+MEETING", text, re.I):
            return "Withdrawn Prior to Meeting"
        if re.search(r"Item\s+Withdrawn|Pages?\s+\d+.*removed.*Item\s+Withdrawn", text, re.I | re.S):
            return "Item Withdrawn"
        if re.search(r"NO\s+MATERIALS", text, re.I):
            return "No Materials"
        if re.search(r"REPORT\s+OF\s+CORRESPONDENCE", text, re.I):
            return "Report of Correspondence"

        # Try up to 3 content pages in case header bleeds across pages
        for page_idx in range(1, min(4, len(doc))):
            page_text = doc[page_idx].get_text()
            title_parts = _parse_title_from_text(page_text)
            if title_parts:
                break

        return " ".join(title_parts).strip()[:280] or "(no title extracted)"

    finally:
        doc.close()


def date_display(date_str: str) -> str:
    y, m, d = date_str.split("-")
    return f"{MONTH_NAMES[int(m)]} {int(d)}, {y}"


def build_manifest(tabs_dir: str = "boepdfs_tabs", out: str = "manifest.json") -> None:
    tabs_path = Path(tabs_dir)
    pdfs = sorted(tabs_path.glob("*.pdf"))

    # Collect: date -> tab_num -> {primary_file, alt_files, title}
    meetings: dict[str, dict[int, dict]] = {}

    for pdf in pdfs:
        m = FILENAME_RE.match(pdf.name)
        if not m:
            continue
        date_str, tab_num_s, src_s = m.group(1), m.group(2), m.group(3)
        tab_num = int(tab_num_s)

        if date_str not in meetings:
            meetings[date_str] = {}
        if tab_num not in meetings[date_str]:
            meetings[date_str][tab_num] = {"primary": None, "alts": [], "title": ""}

        entry = meetings[date_str][tab_num]
        if src_s is None:
            entry["primary"] = pdf.name
        else:
            entry["alts"].append(pdf.name)

    # Extract titles — prefer primary file; fall back to first alt
    total_tabs = 0
    print(f"Extracting titles from {sum(len(v) for v in meetings.values())} tabs...")
    for date_str, tabs in meetings.items():
        for tab_num, entry in tabs.items():
            source_file = entry["primary"] or (entry["alts"][0] if entry["alts"] else None)
            if source_file:
                entry["title"] = extract_title(tabs_path / source_file)
            total_tabs += 1

    # Serialize newest-first
    result_meetings = []
    for date_str in sorted(meetings.keys(), reverse=True):
        tabs_list = []
        for tab_num in sorted(meetings[date_str].keys()):
            e = meetings[date_str][tab_num]
            primary = e["primary"] or e["alts"][0]
            tabs_list.append({
                "num": tab_num,
                "title": e["title"],
                "filename": primary,
                "alt_files": e["alts"] if e["primary"] else e["alts"][1:],
            })
        result_meetings.append({
            "date": date_str,
            "date_display": date_display(date_str),
            "tabs": tabs_list,
        })

    manifest = {
        "generated": str(_date.today()),
        "meetings": result_meetings,
    }

    Path(out).write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(f"Wrote {out}: {len(result_meetings)} meetings, {total_tabs} tabs")


def main():
    parser = argparse.ArgumentParser(description="Build tab manifest JSON.")
    parser.add_argument("--tabs", default="boepdfs_tabs")
    parser.add_argument("--out",  default="manifest.json")
    args = parser.parse_args()
    build_manifest(args.tabs, args.out)


if __name__ == "__main__":
    main()
