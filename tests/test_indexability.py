from keyword_bible import indexability as ix

HTML_NOINDEX = '<html><head><meta name="robots" content="noindex,follow">' \
    '<link rel="canonical" href="https://x/a/"></head></html>'
HTML_OK = '<html><head><link rel="canonical" href="https://x/a/"></head></html>'

def test_parse_noindex_meta():
    s = ix.parse_signals(HTML_NOINDEX, {}, "https://x/a/", "https://x/a/")
    assert s["noindex"] is True and s["self_canonical"] is True

def test_parse_ok():
    s = ix.parse_signals(HTML_OK, {}, "https://x/a/", "https://x/a/")
    assert s["noindex"] is False and s["canonical"] == "https://x/a/"

def test_x_robots_header():
    s = ix.parse_signals(HTML_OK, {"X-Robots-Tag": "noindex"}, "https://x/a/", "https://x/a/")
    assert s["x_robots_noindex"] is True

def test_verdict_blocked_by_noindex():
    ok, reason = ix.indexable_verdict(200, True, {"noindex": True, "x_robots_noindex": False, "self_canonical": True})
    assert ok is False and "noindex" in reason.lower()

def test_verdict_ok():
    ok, reason = ix.indexable_verdict(200, True, {"noindex": False, "x_robots_noindex": False, "self_canonical": True})
    assert ok is True

def test_gsc_fallback():
    d = ix.gsc_status("https://x/a/")
    assert d["coverage_state"] == "GSC not connected"
