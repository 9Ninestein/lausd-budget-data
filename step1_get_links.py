#!/usr/bin/env python3
"""
STEP 1: Open the LAUSD Treasury page in a browser, find all PDF links,
and save them to links.csv.

Requirements:
    pip install playwright
    python -m playwright install chromium
"""

import csv
import re
import sys
import urllib.parse
from pathlib import Path

PAGE_URL = "https://treasury.lausd.org/apps/pages/index.jsp?uREC_ID=4417350&type=d&pREC_ID=2650401"
BASE_URL  = "https://treasury.lausd.org"
OUTPUT_CSV = Path("links.csv")


def main():
    from playwright.sync_api import sync_playwright

    print("Opening browser...")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(PAGE_URL, timeout=60000)

        print()
        print("A browser window has opened.")
        print("Wait for the page to fully load (solve any CAPTCHA if prompted).")
        input("Then press Enter here to collect the links...")
        print()

        # Grab all <a href> tags from every frame on the page
        links = []
        for frame in page.frames:
            try:
                anchors = frame.eval_on_selector_all(
                    "a[href]",
                    "els => els.map(el => ({ href: el.href, text: el.innerText.trim() }))"
                )
                links.extend(anchors)
            except Exception:
                pass

        browser.close()

    # Filter to PDF links only — deduplicate by URL but keep every unique URL
    pdf_links = []
    seen_urls = set()
    for link in links:
        href = link.get("href", "")
        if not re.search(r"\.pdf(\?|#|$)", href, re.IGNORECASE):
            continue
        full_url = urllib.parse.urljoin(BASE_URL, href)
        if full_url in seen_urls:
            continue
        seen_urls.add(full_url)
        text = re.sub(r"\s+", " ", link.get("text", "")).strip()
        pdf_links.append({"url": full_url, "text": text})
        print(f"  Found: {text or '(no text)'} → {full_url}")

    if not pdf_links:
        print("No PDF links found. Check that the page loaded correctly and try again.")
        sys.exit(1)

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["url", "text"])
        writer.writeheader()
        writer.writerows(pdf_links)

    print(f"Found {len(pdf_links)} PDF link(s). Saved to {OUTPUT_CSV}.")
    print(f"Open {OUTPUT_CSV} to verify the links, then run step2_download.py.")


if __name__ == "__main__":
    main()
