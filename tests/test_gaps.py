from keyword_bible import gaps

def test_primary_intent_priority():
    assert gaps.primary_intent(["informational","commercial"]) == "commercial"
    assert gaps.primary_intent([]) == "informational"

def test_missing():
    assert gaps.gap_tier(None, [3, 12]) == "Missing"

def test_winning_only_pb():
    assert gaps.gap_tier(4, [None, None]) == "Winning"

def test_winning_pb_best():
    assert gaps.gap_tier(2, [5, 9]) == "Winning"

def test_near():
    assert gaps.gap_tier(8, [3]) == "Near"

def test_weak():
    assert gaps.gap_tier(45, [6]) == "Weak"

def test_opportunity_formula():
    # 4400 * Missing(1.0) * commercial(1.15) = 5060
    assert gaps.opportunity(4400, "Missing", ["commercial"]) == 5060
    assert gaps.opportunity(4400, "Winning", ["commercial"]) == 0
