import re
from urllib.parse import urlparse
from . import config

_SUFFIXES = {"hiart", "hipvd", "tipvd", "tic"}
_STOP = {"the", "a", "an", "of", "in", "for", "to", "and", "insurance", "ae", "en", "ar"}

def _tokens(text):
    return {t for t in re.split(r"[^a-z0-9؀-ۿ]+", text.lower()) if t and t not in _STOP}

def slug_tokens(url):
    path = urlparse(url).path
    toks = _tokens(path)
    return {t for t in toks if t not in _SUFFIXES}

def _jaccard(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)

def suggest(keyword_norm, pb_ranking_url, sitemap_urls, sitemap_set):
    if pb_ranking_url:
        if pb_ranking_url in sitemap_set:
            return (pb_ranking_url, "current")
        return (pb_ranking_url, "not-in-sitemap")
    kw = _tokens(keyword_norm)
    best_url, best_score = None, 0.0
    for u in sitemap_urls:
        score = _jaccard(kw, slug_tokens(u))
        if score > best_score:
            best_url, best_score = u, score
    if best_score >= config.URL_MATCH_THRESHOLD:
        status = "matched" if best_score >= config.URL_LOWCONF_THRESHOLD else "low-confidence — verify"
        return (best_url, status)
    return ("new page needed", "new-page")
