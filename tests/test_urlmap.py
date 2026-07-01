from keyword_bible import urlmap

SM = [
    "https://www.policybazaar.ae/health-insurance/",
    "https://www.policybazaar.ae/daman-health-insurance-hipvd/",
    "https://www.policybazaar.ae/schengen-travel-insurance-tic/",
]
SMSET = set(SM)

def test_slug_tokens_drop_suffix():
    toks = urlmap.slug_tokens("https://www.policybazaar.ae/daman-health-insurance-hipvd/")
    assert "daman" in toks and "health" in toks and "hipvd" not in toks

def test_pb_ranks_and_in_sitemap_is_current():
    url, status = urlmap.suggest("health insurance", "https://www.policybazaar.ae/health-insurance/", SM, SMSET)
    assert status == "current" and url == "https://www.policybazaar.ae/health-insurance/"

def test_pb_ranks_but_url_not_in_sitemap():
    url, status = urlmap.suggest("health insurance", "https://www.policybazaar.ae/old-page/", SM, SMSET)
    assert status == "not-in-sitemap"

def test_missing_keyword_matches_best_url():
    url, status = urlmap.suggest("daman health insurance", "", SM, SMSET)
    assert url == "https://www.policybazaar.ae/daman-health-insurance-hipvd/"
    assert status in ("matched", "low-confidence — verify")

def test_no_match_new_page():
    url, status = urlmap.suggest("zzz qqq nonsense", "", SM, SMSET)
    assert status == "new-page" and url == "new page needed"
