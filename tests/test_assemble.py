from keyword_bible import assemble

def _rec(site, kw, pos, url="", intents=None, vol=1000, serp=None):
    return {"site": site, "keyword": kw, "keyword_norm": kw, "position": pos,
            "volume": vol, "kd": 10.0, "cpc": 1.0, "url": url,
            "intents": intents or ["commercial"], "serp": serp or ["AI overview"]}

def test_build_classifies_and_scores():
    rbs = {
        "policybazaar.ae": {"health insurance dubai": _rec("policybazaar.ae","health insurance dubai",45,"https://www.policybazaar.ae/health-insurance/")},
        "metlife.ae": {"health insurance dubai": _rec("metlife.ae","health insurance dubai",1,"https://www.metlife.ae/x")},
        "shory.com": {"dubai traffic fines": _rec("shory.com","dubai traffic fines",5,"https://www.shory.com/x",["informational"])},
    }
    sm = {"health": ["https://www.policybazaar.ae/health-insurance/"], "travel": []}
    out = assemble.build(rbs, sm)
    health = out["health"]
    assert len(health) == 1
    row = health[0]
    assert row["Keyword"] == "health insurance dubai"
    assert row["PB Position"] == 45 and row["MetLife"] == 1
    assert row["Gap Status"] == "Weak"           # PB 45 > best comp 1
    assert row["Opportunity Score"] > 0
    assert row["AI Overview?"] == "Yes"
    # the traffic-fines keyword is out-of-scope -> not in health/travel/review
    assert out["recon"]["out"] == 1

def test_columns_defined():
    assert "Gap Status" in assemble.MATRIX_COLUMNS
    assert assemble.GAP_COLUMNS[0] == "Keyword"
