#!/usr/bin/env python3
"""
Download all PDFs linked on the LAUSD Treasury budget page.
URL: https://treasury.lausd.org/apps/pages/index.jsp?uREC_ID=4417350&type=d&pREC_ID=2650401

Requirements:
    pip install playwright beautifulsoup4 requests
    playwright install chromium

If playwright is unavailable, falls back to requests (may be blocked by WAF on cloud IPs).
"""

import os
import re
import sys
import time
import urllib.parse
from pathlib import Path

PAGE_URL = "https://treasury.lausd.org/apps/pages/index.jsp?uREC_ID=4417350&type=d&pREC_ID=2650401"
BASE_URL = "https://treasury.lausd.org"
OUTPUT_DIR = Path("pdfs")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return name or "document.pdf"


def pdf_links_from_html(html: str) -> list[tuple[str, str]]:
    """Parse HTML and return (url, filename) pairs for all PDF links."""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    results = []
    seen: set[str] = set()

    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if not (re.search(r"\.pdf(\?|#|$)", href, re.IGNORECASE) or "/pdf/" in href.lower()):
            continue
        full_url = urllib.parse.urljoin(BASE_URL, href)
        if full_url in seen:
            continue
        seen.add(full_url)

        url_path = urllib.parse.unquote(urllib.parse.urlparse(full_url).path)
        url_filename = url_path.rsplit("/", 1)[-1]
        link_text = tag.get_text(strip=True)
        if url_filename.lower().endswith(".pdf"):
            filename = sanitize_filename(url_filename)
        elif link_text:
            filename = sanitize_filename(link_text + ".pdf")
        else:
            filename = f"document_{len(results) + 1}.pdf"

        results.append((full_url, filename))

    return results


# ---------------------------------------------------------------------------
# Strategy 1: Playwright (real Chromium browser — bypasses most WAFs)
# ---------------------------------------------------------------------------

def fetch_with_playwright() -> list[tuple[str, str]]:
    from playwright.sync_api import sync_playwright

    print("Using Playwright (Chromium) to fetch the page...")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="en-US",
        )
        page = context.new_page()
        page.goto(PAGE_URL, wait_until="networkidle", timeout=30000)
        html = page.content()
        browser.close()

    return pdf_links_from_html(html)


# ---------------------------------------------------------------------------
# Strategy 2: requests + BeautifulSoup (works if IP is not blocked)
# ---------------------------------------------------------------------------

def fetch_with_requests() -> list[tuple[str, str]]:
    import requests

    print("Using requests to fetch the page...")
    session = requests.Session()
    # Warm up with base domain to collect cookies
    try:
        session.get(BASE_URL, headers=HEADERS, timeout=15)
        time.sleep(0.5)
    except Exception:
        pass

    response = session.get(PAGE_URL, headers={**HEADERS, "Referer": BASE_URL}, timeout=30)
    response.raise_for_status()
    return pdf_links_from_html(response.text)


# ---------------------------------------------------------------------------
# Downloader
# ---------------------------------------------------------------------------

def download_pdf(url: str, dest: Path) -> bool:
    import requests

    tmp = dest.with_suffix(".tmp")
    try:
        session = requests.Session()
        response = session.get(url, headers=HEADERS, timeout=60, stream=True)
        response.raise_for_status()

        with open(tmp, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        # Verify the file is actually a PDF (starts with %PDF magic bytes)
        with open(tmp, "rb") as f:
            header = f.read(5)
        if header != b"%PDF-":
            tmp.unlink(missing_ok=True)
            print(f"  ERROR: server did not return a valid PDF (got HTML/error page instead)")
            print(f"  This URL may require a logged-in session or has additional access restrictions.")
            return False

        tmp.rename(dest)
        size_kb = dest.stat().st_size / 1024
        print(f"  Saved {dest.name} ({size_kb:.1f} KB)")
        return True
    except Exception as e:
        tmp.unlink(missing_ok=True)
        print(f"  ERROR: {e}")
        return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Try Playwright first, fall back to requests
    pdf_links: list[tuple[str, str]] = []
    try:
        import playwright  # noqa: F401
        pdf_links = fetch_with_playwright()
    except ImportError:
        print("playwright not installed; falling back to requests.")
        print("  To install: pip install playwright && playwright install chromium\n")
        try:
            pdf_links = fetch_with_requests()
        except Exception as e:
            print(f"requests also failed: {e}")
            print("\nThe site appears to block automated access from this IP.")
            print("Run this script from your local machine (not a cloud server).")
            sys.exit(1)
    except Exception as e:
        print(f"Playwright failed ({e}); falling back to requests.")
        try:
            pdf_links = fetch_with_requests()
        except Exception as e2:
            print(f"requests also failed: {e2}")
            print("\nThe site appears to block automated access from this IP.")
            print("Run this script from your local machine (not a cloud server).")
            sys.exit(1)

    if not pdf_links:
        print("No PDF links found on the page.")
        sys.exit(1)

    print(f"\nFound {len(pdf_links)} PDF(s). Downloading to ./{OUTPUT_DIR}/\n")

    success, failed = 0, 0
    for i, (url, filename) in enumerate(pdf_links, 1):
        dest = OUTPUT_DIR / filename
        if dest.exists():
            # Skip only if it's a valid PDF; re-download if it was a bad file
            with open(dest, "rb") as f:
                header = f.read(5)
            if header == b"%PDF-":
                print(f"[{i}/{len(pdf_links)}] Skipping (exists): {filename}")
                success += 1
                continue
            else:
                print(f"[{i}/{len(pdf_links)}] Re-downloading (previously corrupt): {filename}")
                dest.unlink()
        print(f"[{i}/{len(pdf_links)}] {filename}")
        print(f"  {url}")
        if download_pdf(url, dest):
            success += 1
        else:
            failed += 1
        if i < len(pdf_links):
            time.sleep(0.5)

    print(f"\nDone. {success} downloaded, {failed} failed.")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
