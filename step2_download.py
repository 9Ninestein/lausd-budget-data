#!/usr/bin/env python3
"""
STEP 2: Read a links CSV and download each PDF.

Usage:
    python step2_download.py                        # reads links.csv (default)
    python step2_download.py --csv boe_links.csv    # reads a different CSV

Requirements:
    pip install requests
"""

import argparse
import csv
import re
import sys
import time
import random
from pathlib import Path

import requests

INPUT_CSV  = Path("links.csv")
OUTPUT_DIR = Path("pdfs")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
}


def sanitize(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return name if name else "document"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", default=str(INPUT_CSV), help="Path to the links CSV file")
    args = parser.parse_args()
    input_csv = Path(args.csv)

    if not input_csv.exists():
        print(f"{input_csv} not found. Run the appropriate step1 script first.")
        sys.exit(1)

    with open(input_csv, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        print(f"{INPUT_CSV} is empty.")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"Downloading {len(rows)} PDF(s) to ./{OUTPUT_DIR}/\n")

    session = requests.Session()
    success, failed = 0, 0
    used_filenames: set[str] = set()  # track names assigned this run to avoid collisions

    for i, row in enumerate(rows, 1):
        url  = row["url"]
        text = row.get("text", "")

        # Prefer the URL filename; fall back to link text; last resort: index
        url_name = url.rsplit("/", 1)[-1].split("?")[0]
        if url_name.lower().endswith(".pdf"):
            base = sanitize(url_name[:-4])  # strip .pdf, re-add below
        elif text:
            base = sanitize(text)
        else:
            base = f"document_{i}"

        # Resolve collisions: if this base name was already used this run,
        # append a counter so nothing gets silently overwritten
        candidate = base + ".pdf"
        counter = 2
        while candidate in used_filenames:
            candidate = f"{base}_({counter}).pdf"
            counter += 1
        filename = candidate
        used_filenames.add(filename)

        dest = OUTPUT_DIR / filename
        tmp  = dest.with_suffix(".tmp")

        if dest.exists():
            with open(dest, "rb") as f:
                if f.read(5) == b"%PDF-":
                    print(f"[{i}/{len(rows)}] Skipping (already downloaded): {filename}")
                    success += 1
                    continue

        print(f"[{i}/{len(rows)}] {filename}")
        print(f"  {url}")

        try:
            response = session.get(url, headers=HEADERS, timeout=60, stream=True)
            response.raise_for_status()

            with open(tmp, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            with open(tmp, "rb") as f:
                magic = f.read(5)

            if magic != b"%PDF-":
                tmp.unlink(missing_ok=True)
                print(f"  ERROR: got an HTML/error page instead of a PDF.")
                failed += 1
            else:
                tmp.rename(dest)
                print(f"  Saved ({dest.stat().st_size // 1024} KB)")
                success += 1

        except Exception as e:
            tmp.unlink(missing_ok=True)
            print(f"  ERROR: {e}")
            failed += 1

        if i < len(rows):
            delay = random.uniform(2, 5)
            print(f"  Waiting {delay:.1f}s...")
            time.sleep(delay)

    print(f"\nDone. {success} saved, {failed} failed.")


if __name__ == "__main__":
    main()
