from keyword_bible import classify

def test_url_path_signal_wins_over_hard_exclude():
    # "car" is hard-excluded, but a /health-insurance/ URL rescues it to health
    v, _ = classify.classify("car accident health cover", ["https://x/health-insurance/foo-hiart/"], ["informational"])
    assert v == "health"

def test_health_word_list():
    assert classify.classify("thiqa renewal", [], ["transactional"])[0] == "health"

def test_travel_word_list():
    assert classify.classify("schengen visa insurance", [], ["commercial"])[0] == "travel"

def test_hard_exclude_out():
    assert classify.classify("dubai traffic fines", [], ["informational"])[0] == "out"

def test_conflict_goes_review():
    assert classify.classify("travel medical health insurance", [], ["commercial"])[0] == "review"

def test_no_signal_goes_review():
    assert classify.classify("best offers 2026", [], ["informational"])[0] == "review"

def test_competitor_brand_navigational_out():
    v, _ = classify.classify("daman login", [], ["navigational"])
    assert v == "out"

def test_pb_brand_kept_not_excluded():
    # PB brand + health word -> health (PB brand not in COMPETITOR_BRANDS)
    v, _ = classify.classify("policybazaar health insurance", [], ["navigational"])
    assert v == "health"

def test_no_substring_false_positive():
    # "art" must not match inside "cart"; "cart offers" has no health/travel signal
    assert classify.classify("cart offers", [], ["informational"])[0] == "review"
