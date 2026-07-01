# tests/test_urlmap_report.py
from keyword_bible import urlmap_report


def fake_index(url):
    return {"http_status": 200, "final_url": url, "robots_allowed": True, "noindex": False,
            "canonical": url, "self_canonical": True, "indexable": True, "reason": "indexable"}


def fake_gsc(url):
    return {"coverage_state": "GSC not connected", "last_crawl": "GSC not connected",
            "google_canonical": "GSC not connected"}


def test_build_rows_counts_keywords_and_indexability():
    sm = ["https://www.policybazaar.ae/health-insurance/"]
    matrix = [
        {"PB Ranking URL": "https://www.policybazaar.ae/health-insurance/", "Search Volume": 4400, "Keyword": "health insurance dubai"},
        {"PB Ranking URL": "https://www.policybazaar.ae/health-insurance/", "Search Volume": 1000, "Keyword": "medical insurance"},
    ]
    rows = urlmap_report.build_rows(sm, "Health", matrix, fake_index, fake_gsc)
    assert len(rows) == 1
    r = rows[0]
    assert r["# Keywords"] == 2
    assert r["Live: HTTP Status"] == 200 and r["Live: Indexable"] is True
    assert r["GSC: Coverage State"] == "GSC not connected"
    assert r["In Sitemap?"] == "Yes"
