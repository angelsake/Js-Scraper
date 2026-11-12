"""
JavaScript Site Scraper (Playwright) & Excel Report Generator
-------------------------------------------------------------
Many modern websites load their content with JavaScript, so a plain
requests + BeautifulSoup scraper sees an empty page. This tool drives a real
(headless) browser with Playwright, waits for the content to render, collects
every page, cleans the data with pandas, and exports an Excel report.

Practice target (built for learning scraping): https://quotes.toscrape.com/js/

Usage:
    pip install -r requirements.txt
    playwright install chromium
    python scraper.py                  # scrape everything
    python scraper.py --compare        # also show why plain requests fails
    python scraper.py --max-pages 3    # quick test run
"""

import argparse
import time
from datetime import datetime, timezone
from urllib.parse import urljoin

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

START_URL = "https://quotes.toscrape.com/js/"
OUTPUT_FILE = "quotes_report.xlsx"
PAGE_TIMEOUT_MS = 30_000     # max time to load a page
RENDER_TIMEOUT_MS = 15_000   # max time to wait for JavaScript to draw the content
MAX_RETRIES = 3              # bounded retries per page
BACKOFF_SECONDS = 1          # wait grows with each retry (1s, 2s, 3s)
DELAY_SECONDS = 0.5          # be polite: pause between pages

# Runs inside the browser, after JavaScript has drawn the quotes.
EXTRACT_JS = """
els => els.map(el => ({
    text: el.querySelector('span.text').innerText,
    author: el.querySelector('small.author').innerText,
    author_url: (el.querySelector('a[href^="/author/"]') || {getAttribute: () => ''}).getAttribute('href'),
    tags: Array.from(el.querySelectorAll('a.tag')).map(t => t.innerText)
}))
"""


# --------------------------------------------------------------------------
# Scraping
# --------------------------------------------------------------------------
def scrape_page(page, url):
    """Load one page, wait for JavaScript to render, return (quotes, next_url)."""
    page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
    page.wait_for_selector("div.quote", timeout=RENDER_TIMEOUT_MS)
    quotes = page.eval_on_selector_all("div.quote", EXTRACT_JS)

    for quote in quotes:
        href = quote.get("author_url") or ""
        quote["author_url"] = urljoin(page.url, href) if href else ""

    next_href = None
    if page.query_selector("li.next a"):
        next_href = page.eval_on_selector("li.next a", "el => el.getAttribute('href')")
    next_url = urljoin(page.url, next_href) if next_href else None
    return quotes, next_url


def scrape_with_retries(page, url, retries=MAX_RETRIES):
    """Try a page a limited number of times. Returns (result, error_message)."""
    last_error = ""
    for attempt in range(1, retries + 1):
        try:
            return scrape_page(page, url), ""
        except Exception as error:  # timeouts, navigation errors, missing content
            last_error = f"{type(error).__name__}: {str(error).splitlines()[0]}"
            print(f"  Attempt {attempt}/{retries} failed: {last_error}")
            time.sleep(BACKOFF_SECONDS * attempt)
    return None, last_error


def scrape_all(page, start_url=START_URL, max_pages=None):
    """
    Follow the 'Next' button until the last page.
    Never reports a partial run as complete: failures and early stops are labelled.
    """
    records, failures = [], []
    url, page_number, pages_ok = start_url, 1, 0

    while url and (max_pages is None or page_number <= max_pages):
        print(f"Scraping page {page_number}: {url}")
        result, error = scrape_with_retries(page, url)
        if result is None:
            failures.append({"page_number": page_number, "url": url, "error": error})
            url = None
            break

        quotes, next_url = result
        scraped_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        for quote in quotes:
            quote["page_number"] = page_number
            quote["source_url"] = url
            quote["scraped_at"] = scraped_at
        records.extend(quotes)
        pages_ok += 1
        page_number += 1
        url = next_url
        time.sleep(DELAY_SECONDS)

    if failures:
        status = "PARTIAL - a page failed and the crawl stopped there (see Failed Pages)"
    elif url:
        status = "LIMITED - stopped at --max-pages, more pages exist"
    else:
        status = "COMPLETE"

    return {"records": records, "failures": failures, "pages_ok": pages_ok, "status": status}


def compare_static(url):
    """Show why requests + BeautifulSoup is not enough for this site."""
    import requests
    from bs4 import BeautifulSoup

    html = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"}).text
    found = len(BeautifulSoup(html, "html.parser").select("div.quote"))
    print(f"[compare] requests + BeautifulSoup found {found} quote blocks in the raw HTML "
          f"(the browser renders the real ones with JavaScript)\n")


# --------------------------------------------------------------------------
# Cleaning and reports
# --------------------------------------------------------------------------
def clean(records):
    """Return (clean DataFrame, number of duplicates removed)."""
    df = pd.DataFrame(records)

    df["text"] = df["text"].str.strip().str.strip("\u201c\u201d").str.strip()
    df["author"] = df["author"].str.strip()
    df["tag_list"] = df["tags"].apply(lambda tags: [t.strip() for t in tags if t.strip()])
    df["tags"] = df["tag_list"].apply(", ".join)
    df["word_count"] = df["text"].str.split().str.len()

    before = len(df)
    df = df.drop_duplicates(subset=["text", "author"]).reset_index(drop=True)
    return df, before - len(df)


def build_reports(df, duplicates, result, seconds, start_url):
    quotes_df = df[["text", "author", "tags", "word_count", "page_number",
                    "author_url", "source_url", "scraped_at"]]

    authors = (
        df.groupby("author")
        .agg(quotes=("text", "count"), avg_words=("word_count", "mean"))
        .round(1)
        .reset_index()
        .sort_values(["quotes", "author"], ascending=[False, True])
    )

    tags = (
        df["tag_list"].explode()
        .loc[lambda s: s.notna() & (s != "")]
        .value_counts()
        .head(15)
        .rename_axis("tag")
        .reset_index(name="quotes")
    )

    failed = pd.DataFrame(result["failures"], columns=["page_number", "url", "error"])

    summary = pd.DataFrame(
        [
            ("Completion status", result["status"]),
            ("Source", start_url),
            ("Pages scraped successfully", result["pages_ok"]),
            ("Pages failed", len(result["failures"])),
            ("Quotes collected", len(df)),
            ("Duplicates removed", duplicates),
            ("Unique authors", df["author"].nunique()),
            ("Unique tags", df["tag_list"].explode().nunique()),
            ("Average words per quote", round(df["word_count"].mean(), 1)),
            ("Run time (seconds)", round(seconds, 1)),
        ],
        columns=["Metric", "Value"],
    )
    return summary, quotes_df, authors, tags, failed


def style_sheet(ws, widths):
    fill = PatternFill("solid", start_color="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"


def export_excel(path, sheets):
    """sheets: list of (sheet_name, dataframe, column_widths)."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame, widths in sheets:
            frame.to_excel(writer, sheet_name=name, index=False)
            style_sheet(writer.sheets[name], widths)


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def parse_args():
    parser = argparse.ArgumentParser(description="Scrape a JavaScript-rendered site with Playwright.")
    parser.add_argument("--start-url", default=START_URL, help="first page to scrape")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Excel file to create")
    parser.add_argument("--max-pages", type=int, default=None, help="stop after N pages (for testing)")
    parser.add_argument("--headed", action="store_true", help="show the browser window")
    parser.add_argument("--compare", action="store_true",
                        help="show what plain requests + BeautifulSoup would see")
    return parser.parse_args()


def main():
    args = parse_args()
    if args.compare:
        compare_static(args.start_url)

    from playwright.sync_api import sync_playwright  # imported here so tests need no browser

    started = time.time()
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=not args.headed)
        try:
            context = browser.new_context(user_agent="Mozilla/5.0 (portfolio project)")
            page = context.new_page()
            result = scrape_all(page, args.start_url, args.max_pages)
        finally:
            browser.close()

    if not result["records"]:
        print(f"\nNo data collected. Status: {result['status']}")
        for failure in result["failures"]:
            print(f"  Page {failure['page_number']}: {failure['error']}")
        return

    df, duplicates = clean(result["records"])
    summary, quotes_df, authors, tags, failed = build_reports(
        df, duplicates, result, time.time() - started, args.start_url
    )
    export_excel(
        args.output,
        [
            ("Summary", summary, [30, 70]),
            ("Quotes", quotes_df, [90, 24, 40, 12, 12, 45, 45, 24]),
            ("Authors", authors, [28, 10, 12]),
            ("Top Tags", tags, [20, 10]),
            ("Failed Pages", failed, [12, 50, 60]),
        ],
    )
    print(f"\nStatus: {result['status']}")
    print(f"{len(df)} quotes saved to {args.output}")


if __name__ == "__main__":
    main()
