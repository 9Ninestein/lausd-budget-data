#!/usr/bin/env python3
"""
split_tabs.py
=============
Split each board meeting PDF in boepdfs/ into per-tab PDFs.

Each "Tab N" section begins with a page containing only "TAB N" and runs
until the next tab separator (exclusive). The Tab cover page is included.
Output PDFs land in boepdfs_tabs/ and are named like:

    2023-05-09_Tab_03.pdf

When multiple source PDFs share the same meeting date and tab number
(e.g. stamped vs. unstamped versions), a suffix like _src2 is appended
so names never collide.

Usage:
    python split_tabs.py
    python split_tabs.py --indir boepdfs --outdir boepdfs_tabs

Requirements:
    pip install pymupdf
"""

import argparse
import re
import sys
from pathlib import Path

MONTH_MAP = {
    "January": "01", "February": "02", "March": "03", "April": "04",
    "May": "05", "June": "06", "July": "07", "August": "08",
    "September": "09", "October": "10", "November": "11", "December": "12",
}

DATE_RE = re.compile(
    r"\b(January|February|March|April|May|June|July|August"
    r"|September|October|November|December)"
    r"\s+(\d{1,2}),?\s+(\d{4})\b"
)
TAB_LABEL_RE = re.compile(r"\bTab\s+(\d+)\b", re.I)
SOURCE_NUM_RE = re.compile(r"\((\d+)\)")  # extracts N from filename "(N)"


def date_slug(text: str) -> str | None:
    m = DATE_RE.search(text)
    if not m:
        return None
    return f"{m.group(3)}-{MONTH_MAP[m.group(1)]}-{int(m.group(2)):02d}"


def tab_number(page_text: str) -> int | None:
    """Return the tab number if this page is a tab separator, else None.

    Tab separator pages contain 'TAB N' and nothing else meaningful —
    only optional boilerplate like page numbers or 'Return to Order of
    Business'. We detect them by stripping all known boilerplate and
    checking that nothing substantive remains.  This handles every
    ordering variant seen across meeting years:
      - 'TAB 3'
      - 'TAB 3\\nReturn to Order of Business\\n32'
      - '1\\nTAB 1\\nReturn to Order of Business\\n18'
      - 'TAB 2\\n20\\nReturn to Order of Business'
    """
    m = TAB_LABEL_RE.search(page_text)
    if not m:
        return None
    leftover = TAB_LABEL_RE.sub("", page_text)
    leftover = re.sub(r"Return to Order of Business", "", leftover, flags=re.I)
    leftover = re.sub(r"\d+", "", leftover)
    if leftover.strip():
        return None  # real content remains — not a tab separator page
    return int(m.group(1))


def find_tab_pages(doc) -> list[tuple[int, int]]:
    result = []
    for i, page in enumerate(doc):
        n = tab_number(page.get_text().strip())
        if n is not None:
            result.append((i, n))
    return result


def split_pdf(src: Path, outdir: Path, used_names: set) -> int:
    import fitz

    doc = fitz.open(src)
    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        return 0

    slug = date_slug(doc[0].get_text())
    if not slug:
        print(f"  WARNING: no meeting date found on page 1 — skipped")
        doc.close()
        return 0

    tabs = find_tab_pages(doc)
    if not tabs:
        print(f"  no Tab pages — skipped ({total_pages} pages, date {slug})")
        doc.close()
        return 0

    preamble_pages = tabs[0][0]
    if preamble_pages:
        print(f"  {preamble_pages} preamble page(s) before first Tab (not extracted)")

    # Disambiguator from source filename: "(2)" → "_src2", base file → ""
    m = SOURCE_NUM_RE.search(src.stem)
    src_tag = f"_src{m.group(1)}" if m else ""

    saved = 0
    for idx, (page_start, tab_num) in enumerate(tabs):
        page_end = tabs[idx + 1][0] - 1 if idx + 1 < len(tabs) else total_pages - 1

        base_stem = f"{slug}_Tab_{tab_num:02d}"
        candidate = base_stem + ".pdf"
        if candidate in used_names:
            candidate = base_stem + src_tag + ".pdf"
        counter = 2
        while candidate in used_names:
            candidate = base_stem + src_tag + f"_{counter}.pdf"
            counter += 1

        used_names.add(candidate)
        dest = outdir / candidate

        out = fitz.open()
        out.insert_pdf(doc, from_page=page_start, to_page=page_end)
        out.save(dest)
        out.close()
        saved += 1
        print(f"  Tab {tab_num:>2d}  (pp {page_start+1}-{page_end+1})  ->  {candidate}")

    doc.close()
    return saved


def main():
    parser = argparse.ArgumentParser(description="Split BOE meeting PDFs into per-tab files.")
    parser.add_argument("--indir",  default="boepdfs",      help="Folder of source PDFs")
    parser.add_argument("--outdir", default="boepdfs_tabs", help="Folder for tab PDFs")
    args = parser.parse_args()

    indir  = Path(args.indir)
    outdir = Path(args.outdir)

    if not indir.exists():
        print(f"Input folder '{indir}' not found.")
        sys.exit(1)

    pdfs = sorted(indir.glob("*.pdf"))
    if not pdfs:
        print(f"No PDFs found in '{indir}'.")
        sys.exit(1)

    outdir.mkdir(exist_ok=True)
    print(f"Processing {len(pdfs)} PDF(s) from '{indir}' -> '{outdir}'\n")

    used_names: set = set()
    total_tabs = 0
    skipped = 0

    for pdf in pdfs:
        print(pdf.name)
        n = split_pdf(pdf, outdir, used_names)
        if n:
            total_tabs += n
        else:
            skipped += 1
        print()

    print(f"Done. {total_tabs} tab PDF(s) written to '{outdir}/', {skipped} source file(s) skipped.")


if __name__ == "__main__":
    main()
