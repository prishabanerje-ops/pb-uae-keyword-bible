import re
from . import config

def _is_latin(term):
    return all(ord(c) < 128 for c in term)

def contains_term(text, term):
    term = term.lower()
    if " " in term or not _is_latin(term):
        return term in text
    return re.search(r"\b" + re.escape(term) + r"\b", text) is not None

def _matches_any(text, terms):
    return any(contains_term(text, t) for t in terms)

def _url_signal(urls, markers):
    return any(any(m in (u or "").lower() for m in markers) for u in urls)

def classify(keyword_norm, urls, intents):
    text = keyword_norm
    url_health = _url_signal(urls, config.HEALTH_URL_MARKERS)
    url_travel = _url_signal(urls, config.TRAVEL_URL_MARKERS)
    kw_health = _matches_any(text, config.HEALTH_SEEDS)
    kw_travel = _matches_any(text, config.TRAVEL_SEEDS)

    health = url_health or kw_health
    travel = url_travel or kw_travel

    # Competitor-brand navigational exclusion (unless a vertical URL rescues it)
    intents_l = [i.lower() for i in intents]
    if "navigational" in intents_l and _matches_any(text, config.COMPETITOR_BRANDS) \
            and not (url_health or url_travel):
        return ("out", "competitor-brand-navigational")

    if health and travel:
        return ("review", "conflict:health+travel")
    if health:
        return ("health", "url" if url_health else "wordlist")
    if travel:
        return ("travel", "url" if url_travel else "wordlist")
    if _matches_any(text, config.HARD_EXCLUDE):
        return ("out", "hard-exclude")
    return ("review", "no-signal")
