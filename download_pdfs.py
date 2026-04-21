#!/usr/bin/env python3
"""
Download all PDFs linked on the LAUSD Treasury budget page.
URL: https://treasury.lausd.org/apps/pages/index.jsp?uREC_ID=4417350&type=d&pREC_ID=2650401

Requirements:
    pip install playwright beautifulsoup4 requests
    python -m playwright install chromium
"""

import random
import re
import sys
import time
import urllib.parse
from pathlib import Path

PAGE_URL = "https://treasury.lausd.org/apps/pages/index.jsp?uREC_ID=4417350&type=d&pREC_ID=2650401"
BASE_URL = "https://treasury.lausd.org"
OUTPUT_DIR = Path("pdfs")


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


def run(output_dir: Path) -> None:
    from playwright.sync_api import sync_playwright

    output_dir.mkdir(exist_ok=True)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=False,
            slow_mo=80,  # slow every browser action by 80ms to appear more human
        )
        context = browser.new_context(accept_downloads=True)

        # Step 1: load the index page
        print("Opening index page in browser window...")
        print("If you see a Cloudflare challenge or CAPTCHA, solve it manually.")
        print("The script will wait for you.\n")
        page = context.new_page()
        page.goto(PAGE_URL, wait_until="networkidle", timeout=60000)

        # Scroll slowly down the page to mimic a human reading it
        page.mouse.wheel(0, 300)
        time.sleep(2)
        page.mouse.wheel(0, 300)
        time.sleep(1)

        # Pause so the user can confirm the page looks right / solve any CAPTCHA
        input("Press Enter once the page has fully loaded in the browser window...")

        # Collect HTML from the main page and all iframes
        frames = [page.main_frame] + page.frames
        all_html = []
        for frame in frames:
            try:
                all_html.append(frame.content())
            except Exception:
                pass
        combined_html = "\n".join(all_html)

        # Save debug HTML for inspection
        debug_file = Path("page_debug.html")
        debug_file.write_text(combined_html, encoding="utf-8")
        print(f"  (Page HTML saved to {debug_file} for inspection)")

        pdf_links = pdf_links_from_html(combined_html)
        if not pdf_links:
            print("No PDF links found on the page.")
            print(f"Open {debug_file} in a text editor and search for '.pdf' to see what's there.")
            browser.close()
            sys.exit(1)

        print(f"Found {len(pdf_links)} PDF(s). Downloading to ./{output_dir}/\n")

        success, failed = 0, 0
        for i, (url, filename) in enumerate(pdf_links, 1):
            dest = output_dir / filename

            # Skip already-downloaded valid PDFs
            if dest.exists():
                with open(dest, "rb") as f:
                    magic = f.read(5)
                if magic == b"%PDF-":
                    print(f"[{i}/{len(pdf_links)}] Skipping (exists): {filename}")
                    success += 1
                    continue
                else:
                    print(f"[{i}/{len(pdf_links)}] Re-downloading (previously corrupt): {filename}")
                    dest.unlink()

            print(f"[{i}/{len(pdf_links)}] {filename}")
            print(f"  {url}")

            tmp = dest.with_suffix(".tmp")
            try:
                # Navigate to the PDF URL inside the same browser session.
                # Intercept the response to capture the raw bytes — this keeps
                # all session cookies and browser headers intact, which is why
                # it succeeds where a plain HTTP download would be blocked.
                pdf_bytes: bytes | None = None

                def handle_response(response):
                    nonlocal pdf_bytes
                    if response.url == url and response.status == 200:
                        try:
                            pdf_bytes = response.body()
                        except Exception:
                            pass

                pdf_page = context.new_page()
                pdf_page.on("response", handle_response)
                pdf_page.goto(url, wait_until="load", timeout=60000)
                pdf_page.close()

                if not pdf_bytes:
                    print(f"  ERROR: no response body captured.")
                    failed += 1
                    continue

                if not pdf_bytes.startswith(b"%PDF-"):
                    print(f"  ERROR: server returned HTML/error page instead of a PDF.")
                    failed += 1
                    continue

                tmp.write_bytes(pdf_bytes)
                tmp.rename(dest)
                print(f"  Saved {dest.name} ({dest.stat().st_size / 1024:.1f} KB)")
                success += 1

            except Exception as e:
                tmp.unlink(missing_ok=True)
                print(f"  ERROR: {e}")
                failed += 1

            if i < len(pdf_links):
                delay = random.uniform(3, 7)
                print(f"  Waiting {delay:.1f}s before next download...")
                time.sleep(delay)

        browser.close()

    print(f"\nDone. {success} downloaded, {failed} failed.")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    run(OUTPUT_DIR)
