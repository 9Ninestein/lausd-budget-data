#!/usr/bin/env python3
"""
STEP 2: Read links.csv produced by step1_get_links.py and download each PDF.

Requirements:
    pip install requests
"""

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
    if not INPUT_CSV.exists():
        print(f"{INPUT_CSV} not found. Run step1_get_links.py first.")
        sys.exit(1)

    with open(INPUT_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    if not rows:
        print(f"{INPUT_CSV} is empty.")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"Downloading {len(rows)} PDF(s) to ./{OUTPUT_DIR}/\n")

    session = requests.Session()
    success, failed = 0, 0

    for i, row in enumerate(rows, 1):
        url  = row["url"]
        text = row.get("text", "")

        # Derive filename from link text, falling back to the URL filename
        url_name = url.rsplit("/", 1)[-1].split("?")[0]
        if text:
            filename = sanitize(text) + ".pdf"
        elif url_name.lower().endswith(".pdf"):
            filename = sanitize(url_name)
        else:
            filename = f"document_{i}.pdf"

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
