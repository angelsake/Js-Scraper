# JavaScript Site Scraper (Playwright) with Excel Report

Many modern websites draw their content with JavaScript. A plain `requests` + BeautifulSoup scraper downloads the page before the JavaScript runs and finds nothing. This tool uses **Playwright** to drive a real headless browser, waits for the content to render, crawls every page, cleans the data, and exports an Excel report.

Practice target (a site built for learning scraping): https://quotes.toscrape.com/js/

## What it does

1. Opens each page in headless Chromium and waits for the JavaScript-rendered content
2. Follows the "Next" button until the last page
3. Extracts quote text, author, author link, tags, page number, source URL and capture time
4. Cleans the data with pandas (trims text, removes typographic quote marks, removes duplicates, counts words)
5. Exports `quotes_report.xlsx` with five sheets:
   - **Summary**: completion status, pages scraped/failed, quotes, duplicates, unique authors and tags, run time
   - **Quotes**: the full dataset
   - **Authors**: quotes per author and average length
   - **Top Tags**: the 15 most common tags
   - **Failed Pages**: every page that failed, with the reason

## Reliability

- **Bounded retries** (3 attempts per page, increasing wait)
- **Honest status:** the Summary shows `COMPLETE`, `PARTIAL` (a page failed) or `LIMITED` (stopped at `--max-pages`). A partial run is never reported as complete
- **Work before a failure is kept**, and the failure is listed with its reason
- **Polite delay** between pages

## Setup and run

```bash
pip install -r requirements.txt
playwright install chromium

python scraper.py                 # scrape everything
python scraper.py --compare       # also show why plain requests fails here
python scraper.py --max-pages 3   # quick test
python scraper.py --headed        # watch the browser work
```

## Tests

```bash
python test_scraper.py
```

8 tests cover the crawl, retries, bounded retries, partial/limited labelling, cleaning, duplicate removal and Excel export. They use a fake page object, so they need no browser or internet. A full run against the live site is done with the commands above.

## Adapting it to another JavaScript site

Change the selectors in `scrape_page()` and the JavaScript in `EXTRACT_JS` (the content selector, the fields, and the "Next" button). The retry, status and Excel code stay the same.

## Tech used

Python, Playwright, pandas, openpyxl, requests, BeautifulSoup

## Need data from a JavaScript-heavy site?

Contact me: *add your Fiverr/Upwork/email link here*
