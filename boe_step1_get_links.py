#!/usr/bin/env python3
"""
STEP 1 (BOE): Crawl the LAUSD Board of Education events page, follow every
"Regular Board Meeting" link, find the "Materials" link on each meeting page,
and save all discovered PDF URLs to boe_links.csv.

Requirements:
    pip install playwright
    python -m playwright install chromium

Run step2_download.py afterwards to download the PDFs.
"""

import csv
import re
import sys
import time
import urllib.parse
from pathlib import Path

EVENTS_URL = "https://boe.lausd.org/apps/events/"
OUTPUT_CSV = Path("boe_links.csv")

DELAY = 1.5  # seconds to wait between page loads


def all_anchors(page) -> list[dict]:
    """Return all <a href> elements from every frame on the current page."""
    results = []
    for frame in page.frames:
        try:
            items = frame.eval_on_selector_all(
                "a[href]",
                "els => els.map(el => ({ href: el.href, text: el.innerText.trim() }))"
            )
            results.extend(items)
        except Exception:
            pass
    return results


def find_links_by_text(page, pattern: str) -> list[dict]:
    """Return anchors whose text matches a regex pattern."""
    return [
        a for a in all_anchors(page)
        if re.search(pattern, a.get("text", ""), re.IGNORECASE)
    ]


def main():
    from playwright.sync_api import sync_playwright

    print("Opening browser...")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # ── Step 1: load the events page ────────────────────────────────────
        page.goto(EVENTS_URL, timeout=60000)
        print()
        print("Browser is open on the events page.")
        print("Wait for it to fully load, then press Enter.")
        input("Press Enter to start collecting links > ")
        print()

        # Collect all "Regular Board Meeting" links.
        # The events list may span multiple pages — keep clicking "Next" if present.
        meeting_links: list[dict] = []
        page_num = 1
        while True:
            meetings = find_links_by_text(page, r"regular board meeting")
            print(f"  Events page {page_num}: found {len(meetings)} Regular Board Meeting link(s)")
            meeting_links.extend(meetings)

            # Look for a "Next" pagination link
            next_links = find_links_by_text(page, r"^\s*next\s*$")
            if not next_links:
                break
            print(f"  Clicking 'Next' to load more events...")
            page.goto(next_links[0]["href"], timeout=30000)
            time.sleep(DELAY)
            page_num += 1

        if not meeting_links:
            print("No 'Regular Board Meeting' links found. Check the page and try again.")
            browser.close()
            sys.exit(1)

        print(f"\nFound {len(meeting_links)} meeting page(s) total. Scanning each for Materials...\n")

        # ── Step 2: visit each meeting page and find the Materials link ──────
        rows = []
        for i, meeting in enumerate(meeting_links, 1):
            meeting_url  = meeting["href"]
            meeting_text = re.sub(r"\s+", " ", meeting["text"]).strip()
            print(f"[{i}/{len(meeting_links)}] {meeting_text}")
            print(f"  {meeting_url}")

            try:
                page.goto(meeting_url, wait_until="networkidle", timeout=30000)
                time.sleep(DELAY)
            except Exception as e:
                print(f"  ERROR loading page: {e}")
                continue

            materials_links = find_links_by_text(page, r"\bmaterials\b")
            if not materials_links:
                print("  No 'Materials' link found — skipping.")
                continue

            for mat in materials_links:
                mat_url  = mat["href"]
                mat_text = re.sub(r"\s+", " ", mat["text"]).strip()
                print(f"  Materials link: {mat_text} → {mat_url}")

                # If the Materials link points directly to a PDF, record it.
                # If it points to another page, follow it and look for PDFs.
                if re.search(r"\.pdf(\?|#|$)", mat_url, re.IGNORECASE):
                    rows.append({
                        "meeting": meeting_text,
                        "meeting_url": meeting_url,
                        "url": mat_url,
                        "text": meeting_text + " - Materials",
                    })
                else:
                    # Follow the Materials page and find PDF links
                    try:
                        page.goto(mat_url, wait_until="networkidle", timeout=30000)
                        time.sleep(DELAY)
                        pdf_anchors = [
                            a for a in all_anchors(page)
                            if re.search(r"\.pdf(\?|#|$)", a.get("href", ""), re.IGNORECASE)
                        ]
                        if pdf_anchors:
                            for a in pdf_anchors:
                                rows.append({
                                    "meeting": meeting_text,
                                    "meeting_url": meeting_url,
                                    "url": a["href"],
                                    "text": re.sub(r"\s+", " ", a.get("text", "")).strip()
                                           or meeting_text + " - Materials",
                                })
                                print(f"    PDF: {a['href']}")
                        else:
                            print(f"  No PDFs found on Materials page.")
                        # Return to the meeting page for the next iteration
                        page.go_back(timeout=15000)
                        time.sleep(DELAY)
                    except Exception as e:
                        print(f"  ERROR on Materials page: {e}")

        browser.close()

    if not rows:
        print("\nNo PDF links found across any meeting pages.")
        sys.exit(1)

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["meeting", "meeting_url", "url", "text"])
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nSaved {len(rows)} PDF link(s) to {OUTPUT_CSV}.")
    print("Open the CSV to verify, then run:  python step2_download.py --csv boe_links.csv")


if __name__ == "__main__":
    main()
