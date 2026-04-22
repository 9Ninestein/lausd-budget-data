#!/usr/bin/env python3
"""
STEP 1 (BOE all months): Visit every monthly events page from a start date
to today, find every "Regular Board Meeting" link, follow it to find the
"Materials" PDF, and save results to boe_links.csv.

URL pattern: https://boe.lausd.org/apps/events/{year}/{month}/?id=0

Usage:
    python boe_all_months_step1.py                      # 2005-01 to today
    python boe_all_months_step1.py --start 2020-01      # custom start

Requirements:
    pip install playwright
    python -m playwright install chromium
"""

import argparse
import csv
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path

OUTPUT_CSV   = Path("boe_links.csv")
BASE_EVENTS  = "https://boe.lausd.org/apps/events/{year}/{month}/?id=0"
DELAY        = 1.5   # seconds between page loads

MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",
    5:"May",6:"June",7:"July",8:"August",
    9:"September",10:"October",11:"November",12:"December",
}


# ── helpers ─────────────────────────────────────────────────────────────────

def all_anchors(page) -> list[dict]:
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


def anchors_matching(page, pattern: str) -> list[dict]:
    return [a for a in all_anchors(page)
            if re.search(pattern, a.get("text", ""), re.IGNORECASE)]


def is_board_meeting_link(text: str) -> bool:
    """Return True if text contains 'regular', 'board', and 'meeting' in any order."""
    t = text.lower()
    return all(word in t for word in ("regular", "board", "meeting"))


def parse_date_from_text(text: str) -> str | None:
    """Try to extract a date string (YYYY-MM-DD) from meeting link text."""
    patterns = [
        (r'(\w+\s+\d{1,2},?\s+\d{4})', ['%B %d, %Y', '%B %d %Y', '%b %d, %Y', '%b %d %Y']),
        (r'(\d{1,2}/\d{1,2}/\d{4})',    ['%m/%d/%Y']),
        (r'(\d{4}-\d{2}-\d{2})',         ['%Y-%m-%d']),
    ]
    for regex, fmts in patterns:
        m = re.search(regex, text)
        if m:
            raw = m.group(1).strip()
            for fmt in fmts:
                try:
                    return datetime.strptime(raw, fmt).strftime('%Y-%m-%d')
                except ValueError:
                    pass
    return None


def months_range(start: date, end: date):
    """Yield (year, month) tuples from start to end inclusive."""
    y, mo = start.year, start.month
    while (y, mo) <= (end.year, end.month):
        yield y, mo
        mo += 1
        if mo > 12:
            mo, y = 1, y + 1


# ── main ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", default="2022-09",
                        help="Earliest year-month to check, e.g. 2022-09 (default: 2022-09)")
    args = parser.parse_args()

    try:
        start_date = datetime.strptime(args.start, "%Y-%m").date().replace(day=1)
    except ValueError:
        print("--start must be in YYYY-MM format, e.g. 2010-06")
        sys.exit(1)

    end_date = date.today().replace(day=1)
    all_months = list(months_range(start_date, end_date))
    print(f"Will scan {len(all_months)} month(s) from {args.start} to "
          f"{end_date.year}-{end_date.month:02d}.\n")

    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page    = context.new_page()

        # Load the first month's page and wait for user confirmation
        first_year, first_month = all_months[0]
        first_url = BASE_EVENTS.format(year=first_year, month=first_month)
        page.goto(first_url, timeout=60000)
        print("Browser is open. Wait for the page to load fully.")
        input("Press Enter to start crawling all months > ")
        print()

        # Open CSV for writing (incremental saves)
        try:
            csv_file = open(OUTPUT_CSV, "w", newline="", encoding="utf-8")
        except PermissionError:
            print(f"\nERROR: Cannot write to {OUTPUT_CSV}.")
            print("It is probably open in Excel. Please close it and try again.")
            browser.close()
            sys.exit(1)
        writer   = csv.DictWriter(
            csv_file,
            fieldnames=["filename", "url", "text", "meeting", "meeting_url"]
        )
        writer.writeheader()
        total_found = 0

        for idx, (year, month) in enumerate(all_months):
            month_label = f"{MONTH_NAMES[month]} {year}"
            url = BASE_EVENTS.format(year=year, month=month)

            # Only navigate if we're not already on this page (first iteration)
            if idx > 0:
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    time.sleep(DELAY)
                except Exception as e:
                    print(f"  [{month_label}] ERROR loading page: {e}")
                    continue

            meetings = [a for a in all_anchors(page) if is_board_meeting_link(a.get("text", ""))]
            if not meetings:
                print(f"  [{month_label}] No Regular Board Meeting links.")
                continue

            print(f"  [{month_label}] {len(meetings)} meeting(s) found.")

            for meeting in meetings:
                meeting_url  = meeting["href"]
                meeting_text = re.sub(r"\s+", " ", meeting["text"]).strip()

                # Derive the date for the filename
                date_str = parse_date_from_text(meeting_text)
                if not date_str:
                    date_str = f"{year}-{month:02d}-00"  # fallback: year-month-unknown

                try:
                    page.goto(meeting_url, wait_until="networkidle", timeout=30000)
                    time.sleep(DELAY)
                except Exception as e:
                    print(f"    ERROR loading meeting page: {e}")
                    page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    time.sleep(DELAY)
                    continue

                materials_links = anchors_matching(page, r"\bmaterials\b")
                if not materials_links:
                    print(f"    {date_str}: no Materials link found.")
                    page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    time.sleep(DELAY)
                    continue

                for mat in materials_links:
                    mat_url = mat["href"]

                    if re.search(r"\.pdf(\?|#|$)", mat_url, re.IGNORECASE):
                        # Direct PDF link
                        filename = f"{date_str}_Board_Meeting_Materials.pdf"
                        writer.writerow({
                            "filename":    filename,
                            "url":         mat_url,
                            "text":        meeting_text,
                            "meeting":     meeting_text,
                            "meeting_url": meeting_url,
                        })
                        csv_file.flush()
                        print(f"    {date_str}: {filename}")
                        total_found += 1
                    else:
                        # Follow Materials page to find PDFs
                        try:
                            page.goto(mat_url, wait_until="networkidle", timeout=30000)
                            time.sleep(DELAY)
                            pdf_anchors = [
                                a for a in all_anchors(page)
                                if re.search(r"\.pdf(\?|#|$)", a.get("href",""), re.IGNORECASE)
                            ]
                            for a in pdf_anchors:
                                filename = f"{date_str}_Board_Meeting_Materials.pdf"
                                writer.writerow({
                                    "filename":    filename,
                                    "url":         a["href"],
                                    "text":        re.sub(r"\s+", " ", a.get("text","")).strip(),
                                    "meeting":     meeting_text,
                                    "meeting_url": meeting_url,
                                })
                                csv_file.flush()
                                print(f"    {date_str}: {filename}")
                                total_found += 1
                            if not pdf_anchors:
                                print(f"    {date_str}: no PDFs on Materials page.")
                            page.goto(meeting_url, wait_until="networkidle", timeout=30000)
                            time.sleep(DELAY)
                        except Exception as e:
                            print(f"    ERROR on Materials page: {e}")

                page.goto(url, wait_until="domcontentloaded", timeout=20000)
                time.sleep(DELAY)

        csv_file.close()
        browser.close()

    print(f"\nDone. {total_found} PDF link(s) saved to {OUTPUT_CSV}.")
    print("Run:  python step2_download.py --csv boe_links.csv --outdir boepdfs")


if __name__ == "__main__":
    main()
