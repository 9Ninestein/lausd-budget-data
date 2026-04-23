#!/usr/bin/env python3
"""
Search downloaded PDFs for pages containing specific text and extract
those pages into new PDFs saved in a results folder.

Usage:
    python search_pdfs.py                         # searches boepdfs/, default text
    python search_pdfs.py --indir pdfs            # search a different folder
    python search_pdfs.py --text "OTHER TEXT"     # search for different text
    python search_pdfs.py --indir boepdfs --outdir my_results

Requirements:
    pip install pymupdf
"""

import argparse
import sys
from pathlib import Path

DEFAULT_TEXT   = "REQUEST FOR APPROVAL OF PROCUREMENT CONTRACTS"
DEFAULT_INDIR  = Path("boepdfs")
DEFAULT_OUTDIR = Path("procurement_pages")


def search_pdf(pdf_path: Path, search_text: str) -> list[int]:
    """Return 0-based page numbers where search_text appears (case-insensitive)."""
    import fitz

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"  ERROR opening file: {e}")
        return []

    matches = []
    needle  = search_text.lower()
    for i, page in enumerate(doc):
        if needle in page.get_text().lower():
            matches.append(i)
    doc.close()
    return matches


def extract_pages(pdf_path: Path, page_nums: list[int], dest: Path) -> None:
    """Save the specified pages from pdf_path into a new PDF at dest."""
    import fitz

    src = fitz.open(pdf_path)
    out = fitz.open()
    for n in page_nums:
        out.insert_pdf(src, from_page=n, to_page=n)
    out.save(dest)
    out.close()
    src.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--indir",  default=str(DEFAULT_INDIR),
                        help=f"Folder of PDFs to search (default: {DEFAULT_INDIR})")
    parser.add_argument("--outdir", default=str(DEFAULT_OUTDIR),
                        help=f"Folder for results (default: {DEFAULT_OUTDIR})")
    parser.add_argument("--text",   default=DEFAULT_TEXT,
                        help="Text to search for (case-insensitive)")
    args = parser.parse_args()

    indir  = Path(args.indir)
    outdir = Path(args.outdir)
    search_text = args.text

    if not indir.exists():
        print(f"Input folder '{indir}' not found.")
        sys.exit(1)

    pdfs = sorted(indir.glob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in '{indir}'.")
        sys.exit(1)

    outdir.mkdir(exist_ok=True)
    print(f"Searching {len(pdfs)} PDF(s) in '{indir}' for:")
    print(f"  \"{search_text}\"\n")

    total_files, total_pages = 0, 0

    for pdf_path in pdfs:
        print(f"{pdf_path.name}")
        matches = search_pdf(pdf_path, search_text)

        if not matches:
            print(f"  (no match)")
            continue

        page_labels = [str(n + 1) for n in matches]  # 1-based for display
        print(f"  Match on page(s): {', '.join(page_labels)}")

        dest = outdir / f"{pdf_path.stem}_procurement.pdf"
        extract_pages(pdf_path, matches, dest)
        print(f"  Saved → {dest.name}")

        total_files += 1
        total_pages += len(matches)

    print(f"\nDone. Found matches in {total_files} file(s), "
          f"{total_pages} page(s) total. Results in '{outdir}/'.")


if __name__ == "__main__":
    main()
