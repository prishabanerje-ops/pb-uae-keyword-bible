from keyword_bible import subtopics


def test_gov_plan():
    assert subtopics.assign("thiqa renewal", "", "health") == "gov-plan"


def test_hospitals_from_url():
    assert subtopics.assign("al qassimi hospital", "https://www.policybazaar.ae/health-insurance/hospitals/al-qassimi/", "health") == "hospitals"


def test_schengen():
    assert subtopics.assign("schengen travel insurance", "", "travel") == "schengen-visa"


def test_core_default_health():
    assert subtopics.assign("health insurance dubai", "", "health") == "core-health-insurance"


def test_core_default_travel():
    assert subtopics.assign("travel insurance uae", "", "travel") == "core-travel-insurance"
