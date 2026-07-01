"""
urlmap_report.py — URL-Map sheet row builder.

Produces one row per PB URL from the sitemap (+ any extra URLs found in matrix
rows) enriched with keyword-count aggregates and live indexability data.
index_fn / gsc_fn are injected so unit-tests need no network.
"""

from collections import defaultdict

URLMAP_COLUMNS = [
    "PB URL", "Vertical", "In Sitemap?", "# Keywords", "Sum Volume", "Top Keyword",
    "Live: HTTP Status", "Live: Robots Allowed", "Live: Noindex", "Live: Canonical",
    "Live: Self-Canonical", "Live: Indexable", "Live: Reason",
    "GSC: Coverage State", "GSC: Last Crawl", "GSC: Google Canonical",
]


def build_rows(sitemap_urls, vertical, matrix_rows, index_fn, gsc_fn):
    """Return one dict per unique URL (sitemap-first, then matrix-only URLs).

    Args:
        sitemap_urls: ordered list of URLs from the sitemap.
        vertical: string label e.g. "Health" or "Travel".
        matrix_rows: list of keyword-matrix row dicts (must have "PB Ranking URL",
                     "Search Volume", "Keyword").
        index_fn: callable(url) -> dict with keys http_status, final_url,
                  robots_allowed, noindex, canonical, self_canonical,
                  indexable, reason.
        gsc_fn: callable(url) -> dict with keys coverage_state, last_crawl,
                google_canonical.

    Returns:
        list[dict] with one entry per URL, columns matching URLMAP_COLUMNS.
    """
    by_url = defaultdict(list)
    for m in matrix_rows:
        u = m.get("PB Ranking URL") or ""
        if u:
            by_url[u].append(m)

    sm_set = set(sitemap_urls)
    # Sitemap URLs first, then any matrix URLs not already in the sitemap
    all_urls = list(dict.fromkeys(list(sitemap_urls) + list(by_url.keys())))

    rows = []
    for u in all_urls:
        kws = by_url.get(u, [])
        top = (
            max(kws, key=lambda m: m.get("Search Volume", 0))["Keyword"]
            if kws else ""
        )
        idx = index_fn(u)
        g = gsc_fn(u)
        rows.append({
            "PB URL": u,
            "Vertical": vertical,
            "In Sitemap?": "Yes" if u in sm_set else "No",
            "# Keywords": len(kws),
            "Sum Volume": sum(m.get("Search Volume", 0) for m in kws),
            "Top Keyword": top,
            "Live: HTTP Status": idx["http_status"],
            "Live: Robots Allowed": idx["robots_allowed"],
            "Live: Noindex": idx["noindex"],
            "Live: Canonical": idx["canonical"],
            "Live: Self-Canonical": idx["self_canonical"],
            "Live: Indexable": idx["indexable"],
            "Live: Reason": idx["reason"],
            "GSC: Coverage State": g["coverage_state"],
            "GSC: Last Crawl": g["last_crawl"],
            "GSC: Google Canonical": g["google_canonical"],
        })
    return rows
