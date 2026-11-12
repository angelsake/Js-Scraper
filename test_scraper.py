"""
Tests for the crawl logic, retries, failure labelling, cleaning and Excel export.
They use a fake page object, so no browser or internet connection is needed.

Run:  python test_scraper.py      (or: pytest)
"""

import os
import tempfile
from urllib.parse import urljoin

import pandas as pd

import scraper

scraper.BACKOFF_SECONDS = 0   # no waiting in tests
scraper.DELAY_SECONDS = 0

BASE = "https://example.test"


def make_quote(n, author="Albert Einstein", tags=("life", "truth")):
    return {
        "text": f"\u201cQuote number {n} is here.\u201d",
        "author": author,
        "author_url": f"/author/{author.replace(' ', '-')}",
        "tags": list(tags),
    }


class FakePage:
    """Behaves like the small part of a Playwright page that the scraper uses."""

    def __init__(self, site, flaky=None, broken=()):
        self.site = site                 # {url: {"quotes": [...], "next": "/page/2/"}}
        self.flaky = dict(flaky or {})   # {url: number of failures before success}
        self.broken = set(broken)        # urls that always fail
        self.url = ""

    def goto(self, url, **kwargs):
        if url in self.broken:
            raise TimeoutError("page.goto: Timeout exceeded\nextra detail")
        if self.flaky.get(url, 0) > 0:
            self.flaky[url] -= 1
            raise TimeoutError("temporary network problem")
        self.url = url

    def wait_for_selector(self, selector, **kwargs):
        if not self.site[self.url]["quotes"]:
            raise TimeoutError("content never rendered")

    def eval_on_selector_all(self, selector, js):
        return [dict(q) for q in self.site[self.url]["quotes"]]

    def query_selector(self, selector):
        return bool(self.site[self.url].get("next"))

    def eval_on_selector(self, selector, js):
        return self.site[self.url]["next"]


def three_page_site():
    return {
        f"{BASE}/page/1/": {"quotes": [make_quote(1), make_quote(2)], "next": "/page/2/"},
        f"{BASE}/page/2/": {"quotes": [make_quote(3, "Jane Austen"), make_quote(4, "Jane Austen")], "next": "/page/3/"},
        f"{BASE}/page/3/": {"quotes": [make_quote(5, "Mark Twain"), make_quote(6, "Mark Twain")], "next": None},
    }


def test_full_crawl_is_complete():
    result = scraper.scrape_all(FakePage(three_page_site()), f"{BASE}/page/1/")
    assert result["status"] == "COMPLETE"
    assert result["pages_ok"] == 3 and len(result["records"]) == 6
    assert result["failures"] == []
    assert result["records"][0]["author_url"] == f"{BASE}/author/Albert-Einstein"
    assert result["records"][0]["source_url"] == f"{BASE}/page/1/"


def test_transient_failure_is_retried():
    page = FakePage(three_page_site(), flaky={f"{BASE}/page/2/": 2})
    result = scraper.scrape_all(page, f"{BASE}/page/1/")
    assert result["status"] == "COMPLETE" and len(result["records"]) == 6


def test_permanent_failure_is_labelled_partial():
    page = FakePage(three_page_site(), broken={f"{BASE}/page/2/"})
    result = scraper.scrape_all(page, f"{BASE}/page/1/")
    assert result["status"].startswith("PARTIAL")
    assert len(result["failures"]) == 1 and result["failures"][0]["page_number"] == 2
    assert "TimeoutError" in result["failures"][0]["error"]
    assert len(result["records"]) == 2          # work before the failure is kept


def test_retries_are_bounded():
    page = FakePage(three_page_site(), flaky={f"{BASE}/page/2/": 99})
    result = scraper.scrape_all(page, f"{BASE}/page/1/")
    assert result["status"].startswith("PARTIAL")
    assert page.flaky[f"{BASE}/page/2/"] == 99 - scraper.MAX_RETRIES


def test_empty_rendered_page_counts_as_failure():
    site = three_page_site()
    site[f"{BASE}/page/2/"]["quotes"] = []
    result = scraper.scrape_all(FakePage(site), f"{BASE}/page/1/")
    assert result["status"].startswith("PARTIAL")


def test_max_pages_is_labelled_limited():
    result = scraper.scrape_all(FakePage(three_page_site()), f"{BASE}/page/1/", max_pages=2)
    assert result["status"].startswith("LIMITED")
    assert result["pages_ok"] == 2 and len(result["records"]) == 4


def test_cleaning_and_duplicates():
    records = [dict(make_quote(1), page_number=1, source_url="u", scraped_at="t"),
               dict(make_quote(1), page_number=2, source_url="u", scraped_at="t"),   # duplicate
               dict(make_quote(2, tags=()), page_number=2, source_url="u", scraped_at="t")]
    df, duplicates = scraper.clean(records)
    assert duplicates == 1 and len(df) == 2
    assert df.loc[0, "text"] == "Quote number 1 is here."       # curly quotes removed
    assert df.loc[0, "tags"] == "life, truth" and df.loc[1, "tags"] == ""
    assert df.loc[0, "word_count"] == 5


def test_excel_export_has_all_sheets():
    result = scraper.scrape_all(FakePage(three_page_site()), f"{BASE}/page/1/")
    df, duplicates = scraper.clean(result["records"])
    reports = scraper.build_reports(df, duplicates, result, 1.5, f"{BASE}/page/1/")
    summary, quotes_df, authors, tags, failed = reports

    with tempfile.TemporaryDirectory() as folder:
        path = os.path.join(folder, "report.xlsx")
        scraper.export_excel(path, [("Summary", summary, [30, 60]), ("Quotes", quotes_df, [60] * 8),
                                    ("Authors", authors, [20, 10, 10]), ("Top Tags", tags, [20, 10]),
                                    ("Failed Pages", failed, [10, 40, 40])])
        sheets = pd.read_excel(path, sheet_name=None)

    assert list(sheets) == ["Summary", "Quotes", "Authors", "Top Tags", "Failed Pages"]
    assert len(sheets["Quotes"]) == 6 and len(sheets["Authors"]) == 3
    assert sheets["Summary"].iloc[0, 1] == "COMPLETE"
    assert sheets["Top Tags"].iloc[0]["quotes"] == 6


if __name__ == "__main__":
    tests = [value for name, value in sorted(globals().items()) if name.startswith("test_")]
    for test in tests:
        test()
        print(f"PASS  {test.__name__}")
    print(f"\n{len(tests)} tests passed")
