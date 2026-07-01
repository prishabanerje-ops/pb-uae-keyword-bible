import openpyxl
from keyword_bible import ingest

HEADER = ["Keyword","Position","Previous position","Search Volume","Keyword Difficulty",
          "CPC","URL","Traffic","Traffic (%)","Traffic Cost","Competition","Number of Results",
          "Trends","Timestamp","SERP Features by Keyword","Keyword Intents","Position Type"]

def _make_xlsx(path, rows):
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(HEADER)
    for r in rows: ws.append(r)
    wb.save(path)

def test_normalize():
    assert ingest.normalize("  Health   Insurance  ") == "health insurance"

def test_split_multi():
    assert ingest.split_multi("People also ask, Video") == ["People also ask", "Video"]
    assert ingest.split_multi(None) == []

def test_load_site_keeps_best_position(tmp_path):
    p = str(tmp_path/"s.xlsx")
    _make_xlsx(p, [
        ["health insurance dubai", 7, 8, 4400, 42, 1.2, "https://x/health-insurance/", 100, 1, 5, 0.5, 100, "", "", "AI overview", "commercial", "Organic"],
        ["health insurance dubai", 3, 5, 4400, 42, 1.2, "https://x/health-insurance/", 200, 2, 5, 0.5, 100, "", "", "AI overview", "commercial", "Organic"],
    ])
    m = ingest.load_site(p, "x.com")
    assert list(m.keys()) == ["health insurance dubai"]
    rec = m["health insurance dubai"]
    assert rec["position"] == 3 and rec["volume"] == 4400
    assert rec["intents"] == ["commercial"] and rec["serp"] == ["AI overview"]
