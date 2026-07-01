# Health & Travel Keyword Bible — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single, auditable Excel "Keyword Bible" for PolicyBazaar UAE covering Health + Travel keyword mapping, URL mapping, tiered gaps, and URL indexability (Goals G1–G4), from 9 SEMrush exports + 2 sitemaps + GSC.

**Architecture:** A deterministic Python package. Pure-logic modules (ingest → classify → subtopics → matrix → gaps → urlmap) transform the SEMrush rows into per-keyword records; I/O modules (sitemaps, indexability) add live data; an Excel writer emits the 8-sheet workbook; a `build.py` orchestrator wires it end-to-end and runs QA reconciliation. No LLM in the pipeline — every step is scripted and re-runnable.

**Tech Stack:** Python 3.12, `openpyxl` (Excel read/write), Python stdlib only for HTTP (`urllib.request`, `urllib.robotparser`) and XML (`xml.etree`). GSC URL Inspection is an optional adapter with a hard `GSC not connected` fallback. Tests: `pytest`.

## Global Constraints

- **Snapshot date:** 2026-06-29 — copied verbatim into the README sheet; no other date is authoritative.
- **Source files (exact):** the 9 `.xlsx` in `/Users/aakash/Documents/Competator/` (see `config.SEMRUSH_FILES`).
- **SEMrush column order (0-indexed):** 0 Keyword · 1 Position · 2 Previous position · 3 Search Volume · 4 Keyword Difficulty · 5 CPC · 6 URL · 7 Traffic · 8 Traffic(%) · 9 Traffic Cost · 10 Competition · 11 Number of Results · 12 Trends · 13 Timestamp · 14 SERP Features · 15 Keyword Intents · 16 Position Type.
- **Sitemaps (exact):** `https://www.policybazaar.ae/sitemap/sitemap-health-insurance.xml` (1,061 URLs), `https://www.policybazaar.ae/sitemap/sitemap-travel-insurance.xml` (435 URLs).
- **Competitors (exact set, 8):** damaninsurance.ae, sukoon.com, insuranceuae.com (Orient), giggulf.ae (GIG), metlife.ae (MetLife), insurancemarket.ae, shory.com, mediclinic.ae. Home site = policybazaar.ae.
- **No fabrication:** a blank position cell = "absent from that site's top-100 export on 2026-06-29", never a guessed absence. GSC failure = literal `GSC not connected`. Any sitemap/HTTP/GSC failure → stop and report, never infer.
- **Keyword normalisation:** lowercase, strip, collapse internal whitespace. Arabic keywords retained.
- **Verticals in scope:** Health, Travel only. Every input keyword lands in exactly one bucket: `health`, `travel`, `out`, or `review`.
- **HTTP politeness:** custom User-Agent `Mozilla/5.0 (KeywordBibleBot)`, 40s timeout, sequential with a small delay for the indexability crawl.

---

### Task 0: Project scaffold, git init, config

**Files:**
- Create: `/Users/aakash/Documents/Competator/keyword_bible/__init__.py`
- Create: `/Users/aakash/Documents/Competator/keyword_bible/config.py`
- Create: `/Users/aakash/Documents/Competator/requirements.txt`
- Create: `/Users/aakash/Documents/Competator/pytest.ini`
- Test: `/Users/aakash/Documents/Competator/tests/test_config.py`

**Interfaces:**
- Produces: `config.SEMRUSH_FILES: dict[str,str]` (site→abs path), `config.COMPETITORS: list[str]`, `config.HOME: str`, `config.HEALTH_SEEDS/TRAVEL_SEEDS/HARD_EXCLUDE: set[str]`, `config.HEALTH_URL_MARKERS/TRAVEL_URL_MARKERS: list[str]`, `config.COMPETITOR_BRANDS: set[str]`, `config.SEVERITY_WEIGHT/INTENT_MULTIPLIER: dict`, `config.URL_MATCH_THRESHOLD/URL_LOWCONF_THRESHOLD: float`, `config.SITEMAPS: dict[str,str]`, `config.OUT_XLSX: str`, `config.USER_AGENT/HTTP_TIMEOUT`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_config.py
from keyword_bible import config

def test_config_has_nine_semrush_files():
    assert len(config.SEMRUSH_FILES) == 9
    assert config.HOME == "policybazaar.ae"
    assert len(config.COMPETITORS) == 8

def test_seed_lists_nonempty_and_disjoint_verticals():
    assert config.HEALTH_SEEDS and config.TRAVEL_SEEDS and config.HARD_EXCLUDE
    # a term shouldn't be both a core health and core travel seed
    assert config.HEALTH_SEEDS.isdisjoint(config.TRAVEL_SEEDS)

def test_weights_present():
    assert config.SEVERITY_WEIGHT["Missing"] == 1.0
    assert config.INTENT_MULTIPLIER["transactional"] == 1.2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd /Users/aakash/Documents/Competator && python3 -m pytest tests/test_config.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'keyword_bible'`

- [ ] **Step 3: Create the package files**

`keyword_bible/__init__.py`:
```python
# keyword_bible package
```

`keyword_bible/config.py`:
```python
import os

BASE = "/Users/aakash/Documents/Competator"
SNAPSHOT_DATE = "2026-06-29"

HOME = "policybazaar.ae"
COMPETITORS = [
    "damaninsurance.ae", "sukoon.com", "insuranceuae.com", "giggulf.ae",
    "metlife.ae", "insurancemarket.ae", "shory.com", "mediclinic.ae",
]
# Human-readable column labels for competitors (Excel headers)
SITE_LABELS = {
    "policybazaar.ae": "PB", "damaninsurance.ae": "Daman", "sukoon.com": "Sukoon",
    "insuranceuae.com": "Orient", "giggulf.ae": "GIG", "metlife.ae": "MetLife",
    "insurancemarket.ae": "InsuranceMarket", "shory.com": "Shory", "mediclinic.ae": "Mediclinic",
}
# site -> exact filename (verified present in BASE)
SEMRUSH_FILES = {
    "policybazaar.ae": os.path.join(BASE, "policybazaar.ae-organic.Positions-ae-20260629-2026-06-30T07_27_42Z.xlsx"),
    "damaninsurance.ae": os.path.join(BASE, "https___www.damaninsurance.ae_-organic.Positions-ae-20260629-2026-06-30T07_39_31Z.xlsx"),
    "sukoon.com": os.path.join(BASE, "https___www.sukoon.com_-organic.Positions-ae-20260629-2026-06-30T07_41_55Z.xlsx"),
    "insuranceuae.com": os.path.join(BASE, "https___www.insuranceuae.com_-organic.Positions-ae-20260629-2026-06-30T07_40_44Z.xlsx"),
    "giggulf.ae": os.path.join(BASE, "https___www.giggulf.ae_-organic.Positions-ae-20260629-2026-06-30T07_41_05Z.xlsx"),
    "metlife.ae": os.path.join(BASE, "metlife.ae-organic.Positions-ae-20260629-2026-06-30T07_39_18Z.xlsx"),
    "insurancemarket.ae": os.path.join(BASE, "https___insurancemarket.ae_-organic.Positions-ae-20260629-2026-06-30T07_42_32Z.xlsx"),
    "shory.com": os.path.join(BASE, "https___www.shory.com_-organic.Positions-ae-20260629-2026-06-30T07_42_46Z.xlsx"),
    "mediclinic.ae": os.path.join(BASE, "https___www.mediclinic.ae_-organic.Positions-ae-20260629-2026-06-30T07_42_20Z.xlsx"),
}

SITEMAPS = {
    "health": "https://www.policybazaar.ae/sitemap/sitemap-health-insurance.xml",
    "travel": "https://www.policybazaar.ae/sitemap/sitemap-travel-insurance.xml",
}

# URL-path signals (checked against any site's ranking URL, lowercased)
HEALTH_URL_MARKERS = ["/health-insurance", "-hiart", "-hipvd", "/medical", "/hospitals/", "/health/"]
TRAVEL_URL_MARKERS = ["/travel-insurance", "-tipvd", "-tic", "/travel/"]

# Seed word lists (editable; user-approvable). Multi-word seeds match as phrases.
HEALTH_SEEDS = {
    "health insurance", "medical insurance", "health", "medical", "healthcare",
    "thiqa", "daman", "mediclinic", "clinic", "hospital", "dha", "doh", "haad",
    "doctor", "treatment", "medicine", "pharmacy", "dental", "maternity",
    "pregnancy", "ivf", "chronic", "illness", "disease", "copay", "reimbursement",
    "enaya", "saada", "basmah", "isahd", "mednet", "nextcare", "health card",
    "checkup", "vaccination", "family health", "group health", "individual health",
    "تأمين صحي", "ضمان", "صحي", "مستشفى", "علاج", "طبي",
}
TRAVEL_SEEDS = {
    "travel insurance", "travel", "visa", "schengen", "trip", "passport",
    "tourist", "holiday", "vacation", "abroad", "overseas", "baggage", "luggage",
    "immigration", "green list", "quarantine", "single trip", "multi trip",
    "annual travel", "travel medical", "worldwide travel", "international travel",
    "تأمين سفر", "سفر", "تأشيرة", "فيزا",
}
HARD_EXCLUDE = {
    "car", "motor", "vehicle", "mulkiya", "rta", "traffic fine", "salik",
    "tasjeel", "driving license", "bike", "bank", "atm", "balance check",
    "credit card", "loan", "salary", "ratibi", "postal code", "life insurance",
    "home insurance", "pet insurance", "investment", "trading", "savings",
}
# Competitor brand tokens (navigational + these => out; PB brand NOT here so PB brand kept)
COMPETITOR_BRANDS = {
    "daman", "sukoon", "oman insurance", "orient", "oriental", "gig", "metlife",
    "met life", "insurance market", "insurancemarket", "shory", "mediclinic",
}

SEVERITY_WEIGHT = {"Missing": 1.0, "Weak": 0.7, "Near": 0.4, "Winning": 0.0, "None": 0.0}
INTENT_MULTIPLIER = {"transactional": 1.2, "commercial": 1.15, "informational": 1.0, "navigational": 0.6}
INTENT_PRIORITY = ["transactional", "commercial", "informational", "navigational"]

URL_MATCH_THRESHOLD = 0.34      # >= this => accept suggested URL
URL_LOWCONF_THRESHOLD = 0.50    # below this (but >= MATCH) => flag "low-confidence — verify"

OUT_XLSX = os.path.join(BASE, "PolicyBazaar_UAE_Health_Travel_Keyword_Bible.xlsx")
INTERIM_DIR = os.path.join(BASE, "keyword_bible", "_interim")

USER_AGENT = "Mozilla/5.0 (KeywordBibleBot)"
HTTP_TIMEOUT = 40
```

`requirements.txt`:
```
openpyxl>=3.1
pytest>=8.0
```

`pytest.ini`:
```
[pytest]
testpaths = tests
python_files = test_*.py
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/aakash/Documents/Competator && python3 -m pytest tests/test_config.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Init git and commit**

```bash
cd /Users/aakash/Documents/Competator
git init
printf "keyword_bible/_interim/\n__pycache__/\n*.pyc\n.pytest_cache/\n" > .gitignore
git add keyword_bible/__init__.py keyword_bible/config.py requirements.txt pytest.ini tests/test_config.py .gitignore docs/
git commit -m "chore: scaffold keyword_bible package + config"
```

---

### Task 1: Ingest — load SEMrush files into deduped per-(site,keyword) records

**Files:**
- Create: `keyword_bible/ingest.py`
- Test: `tests/test_ingest.py`

**Interfaces:**
- Consumes: `config.SEMRUSH_FILES`.
- Produces:
  - `ingest.normalize(kw: str) -> str`
  - `ingest.split_multi(value) -> list[str]` (splits comma-separated SERP/intent cells; `None`→`[]`)
  - `ingest.load_site(path: str, site: str) -> dict[str, dict]` — key = normalised keyword; value = record dict with keys: `site, keyword, keyword_norm, position:int, volume:int, kd:float, cpc:float, url:str, intents:list[str], serp:list[str]`. Keeps the **best (lowest) position** row per keyword.
  - `ingest.load_all() -> dict[str, dict[str, dict]]` — site → its keyword→record map.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_ingest.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_ingest.py -v`
Expected: FAIL with `AttributeError: module 'keyword_bible.ingest' has no attribute 'normalize'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/ingest.py
import warnings
import openpyxl
from . import config

def normalize(kw):
    return " ".join(str(kw).strip().lower().split())

def split_multi(value):
    if value is None:
        return []
    return [p.strip() for p in str(value).split(",") if p.strip()]

def _to_int(x):
    try: return int(float(x))
    except (TypeError, ValueError): return 0

def _to_float(x):
    try: return float(x)
    except (TypeError, ValueError): return 0.0

def load_site(path, site):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        kw = str(r[0])
        norm = normalize(kw)
        pos = _to_int(r[1])
        if pos <= 0:
            continue
        rec = {
            "site": site, "keyword": kw, "keyword_norm": norm, "position": pos,
            "volume": _to_int(r[3]), "kd": _to_float(r[4]), "cpc": _to_float(r[5]),
            "url": r[6] or "", "intents": split_multi(r[15]), "serp": split_multi(r[14]),
        }
        prev = out.get(norm)
        if prev is None or pos < prev["position"]:
            out[norm] = rec
    wb.close()
    return out

def load_all():
    return {site: load_site(path, site) for site, path in config.SEMRUSH_FILES.items()}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_ingest.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/ingest.py tests/test_ingest.py
git commit -m "feat: ingest SEMrush exports with best-position dedupe"
```

---

### Task 2: Classification — hybrid URL-path + word-list + review

**Files:**
- Create: `keyword_bible/classify.py`
- Test: `tests/test_classify.py`

**Interfaces:**
- Consumes: `config` seed lists/markers/brands.
- Produces:
  - `classify.contains_term(text: str, term: str) -> bool` (word-boundary for latin terms, substring for non-latin/multiword).
  - `classify.classify(keyword_norm: str, urls: list[str], intents: list[str]) -> tuple[str, str]` returning `(vertical, reason)` where vertical ∈ {`health`,`travel`,`out`,`review`}.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_classify.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_classify.py -v`
Expected: FAIL with `AttributeError: module 'keyword_bible.classify' has no attribute 'classify'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/classify.py
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_classify.py -v`
Expected: PASS (9 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/classify.py tests/test_classify.py
git commit -m "feat: hybrid keyword classification (url-path + wordlist + review)"
```

---

### Task 3: Sub-topic clustering

**Files:**
- Create: `keyword_bible/subtopics.py`
- Test: `tests/test_subtopics.py`

**Interfaces:**
- Consumes: nothing external.
- Produces: `subtopics.assign(keyword_norm: str, pb_url: str, vertical: str) -> str`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_subtopics.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_subtopics.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'assign'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/subtopics.py
HEALTH_RULES = [
    ("gov-plan", ["thiqa", "enaya", "saada", "basmah", "isahd", "daman"]),
    ("hospitals", ["hospital", "clinic", "/hospitals/", "medical center", "mediclinic"]),
    ("dental", ["dental", "dentist"]),
    ("maternity", ["maternity", "pregnancy", "ivf"]),
    ("claims-reimbursement", ["claim", "reimburse", "reimbursement"]),
    ("provider-brand", ["mednet", "nextcare", "sukoon", "orient", "metlife"]),
]
TRAVEL_RULES = [
    ("schengen-visa", ["schengen"]),
    ("passport-visa", ["passport", "visa"]),
    ("country-guide", ["green list", "country", "countries", "quarantine"]),
    ("covid-travel", ["covid", "coronavirus"]),
]

def _apply(rules, blob, default):
    for label, needles in rules:
        if any(n in blob for n in needles):
            return label
    return default

def assign(keyword_norm, pb_url, vertical):
    blob = (keyword_norm + " " + (pb_url or "")).lower()
    if vertical == "health":
        return _apply(HEALTH_RULES, blob, "core-health-insurance")
    if vertical == "travel":
        return _apply(TRAVEL_RULES, blob, "core-travel-insurance")
    return "uncategorised"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_subtopics.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/subtopics.py tests/test_subtopics.py
git commit -m "feat: deterministic sub-topic clustering"
```

---

### Task 4: Gap tiers + opportunity score

**Files:**
- Create: `keyword_bible/gaps.py`
- Test: `tests/test_gaps.py`

**Interfaces:**
- Consumes: `config.SEVERITY_WEIGHT`, `config.INTENT_MULTIPLIER`, `config.INTENT_PRIORITY`.
- Produces:
  - `gaps.primary_intent(intents: list[str]) -> str`
  - `gaps.gap_tier(pb_pos: int|None, competitor_positions: list[int|None]) -> str` (`Missing`/`Weak`/`Near`/`Winning`/`None`).
  - `gaps.opportunity(volume: int, tier: str, intents: list[str]) -> int`

Gap definition (exhaustive; refines spec §4.7 to cover all cases): **Missing** = PB absent & ≥1 competitor ranks. **Winning** = PB ranks and (no competitor ranks OR PB ≤ best competitor). **Near** = PB ranks ≤20 and a competitor outranks it. **Weak** = PB ranks >20 and a competitor outranks it. **None** = PB absent & no competitor ranks (shouldn't occur in the in-scope set).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_gaps.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_gaps.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'primary_intent'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/gaps.py
from . import config

def primary_intent(intents):
    il = [i.lower() for i in intents]
    for cand in config.INTENT_PRIORITY:
        if cand in il:
            return cand
    return "informational"

def gap_tier(pb_pos, competitor_positions):
    comps = [p for p in competitor_positions if p is not None]
    best_comp = min(comps) if comps else None
    if pb_pos is None:
        return "Missing" if best_comp is not None else "None"
    if best_comp is None or pb_pos <= best_comp:
        return "Winning"
    return "Near" if pb_pos <= 20 else "Weak"

def opportunity(volume, tier, intents):
    sev = config.SEVERITY_WEIGHT.get(tier, 0.0)
    mult = config.INTENT_MULTIPLIER.get(primary_intent(intents), 1.0)
    return round(volume * sev * mult)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_gaps.py -v`
Expected: PASS (7 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/gaps.py tests/test_gaps.py
git commit -m "feat: gap tiering + opportunity scoring"
```

---

### Task 5: Sitemaps — fetch + parse

**Files:**
- Create: `keyword_bible/sitemaps.py`
- Test: `tests/test_sitemaps.py`

**Interfaces:**
- Consumes: `config.SITEMAPS`, `config.USER_AGENT`, `config.HTTP_TIMEOUT`.
- Produces:
  - `sitemaps.parse(xml_text: str) -> list[str]` (extracts `<loc>` values).
  - `sitemaps.fetch(url: str) -> str` (raises `RuntimeError` on non-200 — stop-and-report).
  - `sitemaps.load_all() -> dict[str, list[str]]` (`health`/`travel` → URL lists).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_sitemaps.py
import pytest
from keyword_bible import sitemaps

XML = """<?xml version='1.0'?>
<urlset xmlns='http://www.sitemaps.org/schemas/sitemap/0.9'>
<url><loc>https://www.policybazaar.ae/health-insurance/</loc></url>
<url><loc>https://www.policybazaar.ae/daman-health-insurance-hipvd/</loc></url>
</urlset>"""

def test_parse_extracts_locs():
    urls = sitemaps.parse(XML)
    assert urls == [
        "https://www.policybazaar.ae/health-insurance/",
        "https://www.policybazaar.ae/daman-health-insurance-hipvd/",
    ]

def test_parse_empty():
    assert sitemaps.parse("<urlset></urlset>") == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_sitemaps.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'parse'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/sitemaps.py
import re
import urllib.request
from . import config

_LOC = re.compile(r"<loc>\s*([^<]+?)\s*</loc>", re.IGNORECASE)

def parse(xml_text):
    return [m.group(1) for m in _LOC.finditer(xml_text)]

def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT})
    with urllib.request.urlopen(req, timeout=config.HTTP_TIMEOUT) as resp:
        if resp.status != 200:
            raise RuntimeError(f"Sitemap fetch failed {resp.status} for {url}")
        return resp.read().decode("utf-8", errors="replace")

def load_all():
    return {vert: parse(fetch(url)) for vert, url in config.SITEMAPS.items()}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_sitemaps.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/sitemaps.py tests/test_sitemaps.py
git commit -m "feat: sitemap fetch + parse"
```

---

### Task 6: URL mapping — suggested URL via token overlap

**Files:**
- Create: `keyword_bible/urlmap.py`
- Test: `tests/test_urlmap.py`

**Interfaces:**
- Consumes: `config.URL_MATCH_THRESHOLD`, `config.URL_LOWCONF_THRESHOLD`.
- Produces:
  - `urlmap.slug_tokens(url: str) -> set[str]` (splits path on `/ - _`, drops SEMrush suffixes `hiart/hipvd/tipvd/tic` and stopwords).
  - `urlmap.suggest(keyword_norm: str, pb_ranking_url: str, sitemap_urls: list[str], sitemap_set: set[str]) -> tuple[str, str]` → `(suggested_url_or_flag, status)` where status ∈ {`current`,`not-in-sitemap`,`matched`,`low-confidence — verify`,`new-page`}.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_urlmap.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_urlmap.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'slug_tokens'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/urlmap.py
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_urlmap.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/urlmap.py tests/test_urlmap.py
git commit -m "feat: keyword->URL suggestion via slug token overlap"
```

---

### Task 7: Indexability — live technical check + GSC adapter

**Files:**
- Create: `keyword_bible/indexability.py`
- Test: `tests/test_indexability.py`

**Interfaces:**
- Consumes: `config.USER_AGENT`, `config.HTTP_TIMEOUT`.
- Produces:
  - `indexability.parse_signals(html: str, headers: dict, final_url: str, requested_url: str) -> dict` with keys `noindex:bool`, `canonical:str`, `self_canonical:bool`, `x_robots_noindex:bool`.
  - `indexability.indexable_verdict(status:int, robots_allowed:bool, signals:dict) -> tuple[bool,str]`.
  - `indexability.check_url(url: str) -> dict` (live; keys: `http_status, final_url, robots_allowed, noindex, canonical, self_canonical, indexable, reason`).
  - `indexability.gsc_status(url: str) -> dict` (keys `coverage_state, last_crawl, google_canonical`; returns all `"GSC not connected"` when GSC is unavailable).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_indexability.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_indexability.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'parse_signals'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/indexability.py
import re
import time
import urllib.request
import urllib.robotparser
from urllib.parse import urlparse
from . import config

_META_ROBOTS = re.compile(r'<meta[^>]+name=["\']robots["\'][^>]+content=["\']([^"\']+)["\']', re.I)
_CANON = re.compile(r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)["\']', re.I)

def parse_signals(html, headers, final_url, requested_url):
    meta = _META_ROBOTS.search(html or "")
    noindex = bool(meta and "noindex" in meta.group(1).lower())
    canon_m = _CANON.search(html or "")
    canonical = canon_m.group(1).strip() if canon_m else ""
    xr = headers.get("X-Robots-Tag", "") if headers else ""
    return {
        "noindex": noindex,
        "canonical": canonical,
        "self_canonical": (canonical.rstrip("/") == (requested_url or "").rstrip("/")) if canonical else False,
        "x_robots_noindex": "noindex" in str(xr).lower(),
    }

def indexable_verdict(status, robots_allowed, signals):
    if status != 200:
        return (False, f"http {status}")
    if not robots_allowed:
        return (False, "blocked by robots.txt")
    if signals.get("noindex"):
        return (False, "meta robots noindex")
    if signals.get("x_robots_noindex"):
        return (False, "x-robots-tag noindex")
    return (True, "indexable")

def _robots_allowed(url):
    parts = urlparse(url)
    rp = urllib.robotparser.RobotFileParser()
    rp.set_url(f"{parts.scheme}://{parts.netloc}/robots.txt")
    try:
        rp.read()
    except Exception:
        return True  # if robots.txt unreachable, treat as allowed (documented)
    return rp.can_fetch(config.USER_AGENT, url)

def check_url(url):
    req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT})
    status, final_url, html, headers = 0, url, "", {}
    try:
        with urllib.request.urlopen(req, timeout=config.HTTP_TIMEOUT) as resp:
            status = resp.status
            final_url = resp.geturl()
            headers = dict(resp.headers)
            html = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        status = e.code
        headers = dict(e.headers or {})
    except Exception as e:
        return {"http_status": "ERROR", "final_url": url, "robots_allowed": "",
                "noindex": "", "canonical": "", "self_canonical": "",
                "indexable": False, "reason": f"fetch error: {e}"}
    allowed = _robots_allowed(url)
    sig = parse_signals(html, headers, final_url, url)
    ok, reason = indexable_verdict(status, allowed, sig)
    time.sleep(0.3)
    return {"http_status": status, "final_url": final_url, "robots_allowed": allowed,
            "noindex": sig["noindex"], "canonical": sig["canonical"],
            "self_canonical": sig["self_canonical"], "indexable": ok, "reason": reason}

def gsc_status(url):
    # GSC URL Inspection adapter. Wired at execution time if credentials exist.
    # Until then, honest fallback (never fabricated).
    return {"coverage_state": "GSC not connected", "last_crawl": "GSC not connected",
            "google_canonical": "GSC not connected"}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_indexability.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/indexability.py tests/test_indexability.py
git commit -m "feat: live indexability check + GSC fallback adapter"
```

---

### Task 8: Excel writer — 8-sheet workbook

**Files:**
- Create: `keyword_bible/excel.py`
- Test: `tests/test_excel.py`

**Interfaces:**
- Consumes: `config.SITE_LABELS`.
- Produces:
  - `excel.write_sheet(wb, title, columns: list[str], rows: list[dict], tier_col: str|None=None)` — writes header + rows, freezes header row, adds autofilter; if `tier_col` given, fills the row's tier cell by gap tier colour.
  - `excel.write_workbook(out_path: str, payload: dict) -> None` — payload keys: `readme_lines:list[str]`, `health_matrix`, `travel_matrix`, `health_gaps`, `travel_gaps`, `health_urlmap`, `travel_urlmap`, `review` (each a `{"columns":[...], "rows":[...]}` dict).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_excel.py
import openpyxl
from keyword_bible import excel

def test_write_workbook_has_eight_sheets(tmp_path):
    block = {"columns": ["Keyword", "Gap Status"], "rows": [{"Keyword": "k", "Gap Status": "Missing"}]}
    payload = {
        "readme_lines": ["Snapshot: 2026-06-29"],
        "health_matrix": block, "travel_matrix": block,
        "health_gaps": block, "travel_gaps": block,
        "health_urlmap": {"columns": ["PB URL"], "rows": [{"PB URL": "u"}]},
        "travel_urlmap": {"columns": ["PB URL"], "rows": [{"PB URL": "u"}]},
        "review": {"columns": ["Keyword", "Reason"], "rows": [{"Keyword": "x", "Reason": "no-signal"}]},
    }
    out = str(tmp_path / "b.xlsx")
    excel.write_workbook(out, payload)
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == [
        "README", "Health — Master Matrix", "Travel — Master Matrix",
        "Health — Gaps", "Travel — Gaps", "Health — URL Map", "Travel — URL Map", "Review",
    ]
    ws = wb["Health — Master Matrix"]
    assert ws["A1"].value == "Keyword" and ws["A2"].value == "k"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_excel.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'write_workbook'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/excel.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TIER_FILL = {
    "Missing": PatternFill("solid", fgColor="F4CCCC"),
    "Weak":    PatternFill("solid", fgColor="FCE5CD"),
    "Near":    PatternFill("solid", fgColor="FFF2CC"),
    "Winning": PatternFill("solid", fgColor="D9EAD3"),
}
HEADER_FILL = PatternFill("solid", fgColor="1155CC")
HEADER_FONT = Font(color="FFFFFF", bold=True)

def write_sheet(wb, title, columns, rows, tier_col=None):
    ws = wb.create_sheet(title=title[:31])
    ws.append(columns)
    for c in ws[1]:
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = Alignment(vertical="center")
    tier_idx = columns.index(tier_col) + 1 if tier_col and tier_col in columns else None
    for r in rows:
        ws.append([r.get(col, "") for col in columns])
        if tier_idx:
            cell = ws.cell(row=ws.max_row, column=tier_idx)
            fill = TIER_FILL.get(cell.value)
            if fill:
                cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(columns, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(len(str(col)) + 2, 12), 60)
    return ws

def write_workbook(out_path, payload):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("README")
    for line in payload["readme_lines"]:
        ws.append([line])
    write_sheet(wb, "Health — Master Matrix", payload["health_matrix"]["columns"], payload["health_matrix"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Travel — Master Matrix", payload["travel_matrix"]["columns"], payload["travel_matrix"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Health — Gaps", payload["health_gaps"]["columns"], payload["health_gaps"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Travel — Gaps", payload["travel_gaps"]["columns"], payload["travel_gaps"]["rows"], tier_col="Gap Status")
    write_sheet(wb, "Health — URL Map", payload["health_urlmap"]["columns"], payload["health_urlmap"]["rows"])
    write_sheet(wb, "Travel — URL Map", payload["travel_urlmap"]["columns"], payload["travel_urlmap"]["rows"])
    write_sheet(wb, "Review", payload["review"]["columns"], payload["review"]["rows"])
    wb.save(out_path)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_excel.py -v`
Expected: PASS (1 passed). Note: sheet titles use an en-dash `—`; keep them identical everywhere.

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/excel.py tests/test_excel.py
git commit -m "feat: 8-sheet Excel writer with tier colour-coding"
```

---

### Task 9: Assembler — build matrix rows, gaps, URL-map rows, review

**Files:**
- Create: `keyword_bible/assemble.py`
- Test: `tests/test_assemble.py`

**Interfaces:**
- Consumes: everything above (`ingest, classify, subtopics, gaps, urlmap, config`).
- Produces:
  - `assemble.build(records_by_site: dict, sitemaps_by_vert: dict) -> dict` returning `{"health":[matrixrow...], "travel":[...], "review":[...], "recon":{...}}` where each matrix row is a dict keyed by the Master-Matrix column names (§5.1 of spec), and `recon` holds bucket counts for QA.
  - `assemble.MATRIX_COLUMNS: list[str]`, `assemble.URLMAP_COLUMNS: list[str]`, `assemble.REVIEW_COLUMNS: list[str]`, `assemble.GAP_COLUMNS: list[str]`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_assemble.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_assemble.py -v`
Expected: FAIL with `AttributeError: ... has no attribute 'build'`

- [ ] **Step 3: Write minimal implementation**

```python
# keyword_bible/assemble.py
from . import config, classify, subtopics, gaps, urlmap

MATRIX_COLUMNS = (
    ["Keyword", "Vertical", "Sub-topic", "Intent", "Search Volume", "Keyword Difficulty",
     "CPC", "PB Position", "PB Ranking URL", "PB Suggested URL"]
    + [config.SITE_LABELS[s] for s in config.COMPETITORS]
    + ["Best Competitor", "Best Competitor Position", "Gap Status", "Opportunity Score",
       "SERP Features", "AI Overview?", "URL Map Status"]
)
GAP_COLUMNS = ["Keyword", "Sub-topic", "Intent", "Search Volume", "PB Position", "Gap Status",
               "Opportunity Score", "Best Competitor", "Best Competitor Position",
               "PB Suggested URL", "URL Map Status"]
URLMAP_COLUMNS = ["PB URL", "Vertical", "In Sitemap?"]  # extended in Task 10 with live/GSC cols
REVIEW_COLUMNS = ["Keyword", "Candidate", "Reason", "Search Volume", "Sample URL"]

def _all_norm_keys(records_by_site):
    keys = set()
    for m in records_by_site.values():
        keys.update(m.keys())
    return keys

def build(records_by_site, sitemaps_by_vert):
    sm_health = sitemaps_by_vert.get("health", [])
    sm_travel = sitemaps_by_vert.get("travel", [])
    sm_all = sm_health + sm_travel
    sm_set = set(sm_all)

    out = {"health": [], "travel": [], "review": [],
           "recon": {"health": 0, "travel": 0, "out": 0, "review": 0, "total": 0}}

    for norm in _all_norm_keys(records_by_site):
        per_site = {s: records_by_site[s].get(norm) for s in records_by_site}
        present = {s: r for s, r in per_site.items() if r}
        urls = [r["url"] for r in present.values() if r["url"]]
        sample = next(iter(present.values()))
        intents = sample["intents"]
        vol = max((r["volume"] for r in present.values()), default=0)
        vertical, reason = classify.classify(norm, urls, intents)
        out["recon"]["total"] += 1

        if vertical in ("out",):
            out["recon"]["out"] += 1
            continue
        if vertical == "review":
            out["recon"]["review"] += 1
            out["review"].append({
                "Keyword": sample["keyword"], "Candidate": reason, "Reason": reason,
                "Search Volume": vol, "Sample URL": urls[0] if urls else "",
            })
            continue

        out["recon"][vertical] += 1
        pb = per_site.get(config.HOME)
        pb_pos = pb["position"] if pb else None
        pb_url = pb["url"] if pb else ""
        comp_positions = [per_site[s]["position"] if per_site.get(s) else None for s in config.COMPETITORS]

        # best competitor
        best_label, best_pos = "", None
        for s in config.COMPETITORS:
            r = per_site.get(s)
            if r and (best_pos is None or r["position"] < best_pos):
                best_pos, best_label = r["position"], config.SITE_LABELS[s]

        tier = gaps.gap_tier(pb_pos, comp_positions)
        opp = gaps.opportunity(vol, tier, intents)
        sub = subtopics.assign(norm, pb_url, vertical)
        sitemap_urls = sm_health if vertical == "health" else sm_travel
        sug_url, sug_status = urlmap.suggest(norm, pb_url, sitemap_urls, sm_set)
        serp = sample["serp"]
        ai = "Yes" if any("ai overview" in s.lower() for s in serp) else "No"

        row = {
            "Keyword": sample["keyword"], "Vertical": vertical.title(), "Sub-topic": sub,
            "Intent": gaps.primary_intent(intents), "Search Volume": vol,
            "Keyword Difficulty": sample["kd"], "CPC": sample["cpc"],
            "PB Position": pb_pos if pb_pos is not None else "",
            "PB Ranking URL": pb_url, "PB Suggested URL": sug_url,
            "Best Competitor": best_label,
            "Best Competitor Position": best_pos if best_pos is not None else "",
            "Gap Status": tier, "Opportunity Score": opp,
            "SERP Features": ", ".join(serp), "AI Overview?": ai, "URL Map Status": sug_status,
        }
        for s in config.COMPETITORS:
            r = per_site.get(s)
            row[config.SITE_LABELS[s]] = r["position"] if r else ""
        out[vertical].append(row)

    out["health"].sort(key=lambda r: r["Opportunity Score"], reverse=True)
    out["travel"].sort(key=lambda r: r["Opportunity Score"], reverse=True)
    return out
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python3 -m pytest tests/test_assemble.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/assemble.py tests/test_assemble.py
git commit -m "feat: assemble matrix rows, gaps, review with QA reconciliation"
```

---

### Task 10: Orchestrator — end-to-end build with live data + QA gate

**Files:**
- Create: `keyword_bible/build.py`
- Create: `keyword_bible/urlmap_report.py` (URL-Map sheet rows from sitemaps + indexability)
- Test: `tests/test_urlmap_report.py`

**Interfaces:**
- Consumes: everything above.
- Produces:
  - `urlmap_report.build_rows(sitemap_urls: list[str], vertical: str, matrix_rows: list[dict], index_fn) -> list[dict]` — one row per PB URL with keyword counts (from matrix rows' `PB Ranking URL`) + indexability via injected `index_fn(url)->dict` (injectable so tests avoid network) + `gsc_fn`.
  - `build.main(run_live: bool = True) -> str` — runs full pipeline, prints reconciliation, writes `config.OUT_XLSX`, dumps interim JSON; returns the output path.

- [ ] **Step 1: Write the failing test** (URL-map report with a fake index function — no network)

```python
# tests/test_urlmap_report.py
from keyword_bible import urlmap_report

def fake_index(url):
    return {"http_status": 200, "final_url": url, "robots_allowed": True, "noindex": False,
            "canonical": url, "self_canonical": True, "indexable": True, "reason": "indexable"}

def fake_gsc(url):
    return {"coverage_state": "GSC not connected", "last_crawl": "GSC not connected",
            "google_canonical": "GSC not connected"}

def test_build_rows_counts_keywords_and_indexability():
    sm = ["https://www.policybazaar.ae/health-insurance/"]
    matrix = [
        {"PB Ranking URL": "https://www.policybazaar.ae/health-insurance/", "Search Volume": 4400, "Keyword": "health insurance dubai"},
        {"PB Ranking URL": "https://www.policybazaar.ae/health-insurance/", "Search Volume": 1000, "Keyword": "medical insurance"},
    ]
    rows = urlmap_report.build_rows(sm, "Health", matrix, fake_index, fake_gsc)
    assert len(rows) == 1
    r = rows[0]
    assert r["# Keywords"] == 2
    assert r["Live: HTTP Status"] == 200 and r["Live: Indexable"] is True
    assert r["GSC: Coverage State"] == "GSC not connected"
    assert r["In Sitemap?"] == "Yes"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python3 -m pytest tests/test_urlmap_report.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'keyword_bible.urlmap_report'`

- [ ] **Step 3: Write minimal implementation**

`keyword_bible/urlmap_report.py`:
```python
from collections import defaultdict

URLMAP_COLUMNS = [
    "PB URL", "Vertical", "In Sitemap?", "# Keywords", "Sum Volume", "Top Keyword",
    "Live: HTTP Status", "Live: Robots Allowed", "Live: Noindex", "Live: Canonical",
    "Live: Self-Canonical", "Live: Indexable", "Live: Reason",
    "GSC: Coverage State", "GSC: Last Crawl", "GSC: Google Canonical",
]

def build_rows(sitemap_urls, vertical, matrix_rows, index_fn, gsc_fn):
    by_url = defaultdict(list)
    for m in matrix_rows:
        u = m.get("PB Ranking URL") or ""
        if u:
            by_url[u].append(m)
    sm_set = set(sitemap_urls)
    all_urls = list(dict.fromkeys(list(sitemap_urls) + list(by_url.keys())))
    rows = []
    for u in all_urls:
        kws = by_url.get(u, [])
        top = max(kws, key=lambda m: m.get("Search Volume", 0))["Keyword"] if kws else ""
        idx = index_fn(u)
        g = gsc_fn(u)
        rows.append({
            "PB URL": u, "Vertical": vertical,
            "In Sitemap?": "Yes" if u in sm_set else "No",
            "# Keywords": len(kws),
            "Sum Volume": sum(m.get("Search Volume", 0) for m in kws),
            "Top Keyword": top,
            "Live: HTTP Status": idx["http_status"], "Live: Robots Allowed": idx["robots_allowed"],
            "Live: Noindex": idx["noindex"], "Live: Canonical": idx["canonical"],
            "Live: Self-Canonical": idx["self_canonical"], "Live: Indexable": idx["indexable"],
            "Live: Reason": idx["reason"],
            "GSC: Coverage State": g["coverage_state"], "GSC: Last Crawl": g["last_crawl"],
            "GSC: Google Canonical": g["google_canonical"],
        })
    return rows
```

`keyword_bible/build.py`:
```python
import json
import os
from . import config, ingest, sitemaps, assemble, excel, indexability, urlmap_report

def _gap_rows(matrix_rows):
    keep = {"Missing", "Weak", "Near"}
    rows = [r for r in matrix_rows if r["Gap Status"] in keep]
    rows.sort(key=lambda r: r["Opportunity Score"], reverse=True)
    return [{c: r.get(c, "") for c in assemble.GAP_COLUMNS} for r in rows]

def main(run_live=True):
    os.makedirs(config.INTERIM_DIR, exist_ok=True)
    print("Loading SEMrush files...")
    rbs = ingest.load_all()
    print("Fetching sitemaps...")
    sm = sitemaps.load_all()
    print(f"  health sitemap: {len(sm['health'])} urls | travel sitemap: {len(sm['travel'])} urls")

    data = assemble.build(rbs, sm)
    recon = data["recon"]
    print(f"RECONCILIATION: total={recon['total']} health={recon['health']} "
          f"travel={recon['travel']} out={recon['out']} review={recon['review']}")
    assert recon["health"] + recon["travel"] + recon["out"] + recon["review"] == recon["total"], \
        "Reconciliation failed — bucket counts do not sum to total"

    index_fn = indexability.check_url if run_live else (lambda u: {
        "http_status": "SKIPPED", "final_url": u, "robots_allowed": "", "noindex": "",
        "canonical": "", "self_canonical": "", "indexable": "", "reason": "live check skipped"})
    gsc_fn = indexability.gsc_status

    print("Building URL-Map (live indexability)..." if run_live else "Building URL-Map (skipped live)...")
    health_urlmap = urlmap_report.build_rows(sm["health"], "Health", data["health"], index_fn, gsc_fn)
    travel_urlmap = urlmap_report.build_rows(sm["travel"], "Travel", data["travel"], index_fn, gsc_fn)

    readme = [
        "PolicyBazaar UAE — Health & Travel Keyword Bible",
        f"Snapshot date: {config.SNAPSHOT_DATE}",
        f"Reconciliation: total={recon['total']} health={recon['health']} travel={recon['travel']} out={recon['out']} review={recon['review']}",
        "Gap tiers: Missing / Weak / Near / Winning. Opportunity = Volume x severity x intent multiplier.",
        "IMPORTANT: a blank competitor position = 'absent from that site's top-100 SEMrush export on the snapshot date', NOT a guaranteed non-ranking.",
        "GSC columns show 'GSC not connected' unless GSC URL Inspection was wired at build time.",
        "Competitors: " + ", ".join(config.SITE_LABELS[s] for s in config.COMPETITORS),
    ]
    payload = {
        "readme_lines": readme,
        "health_matrix": {"columns": assemble.MATRIX_COLUMNS, "rows": data["health"]},
        "travel_matrix": {"columns": assemble.MATRIX_COLUMNS, "rows": data["travel"]},
        "health_gaps": {"columns": assemble.GAP_COLUMNS, "rows": _gap_rows(data["health"])},
        "travel_gaps": {"columns": assemble.GAP_COLUMNS, "rows": _gap_rows(data["travel"])},
        "health_urlmap": {"columns": urlmap_report.URLMAP_COLUMNS, "rows": health_urlmap},
        "travel_urlmap": {"columns": urlmap_report.URLMAP_COLUMNS, "rows": travel_urlmap},
        "review": {"columns": assemble.REVIEW_COLUMNS, "rows": data["review"]},
    }
    with open(os.path.join(config.INTERIM_DIR, "assembled.json"), "w") as f:
        json.dump({k: v for k, v in data.items() if k != "recon"}, f, ensure_ascii=False)
    excel.write_workbook(config.OUT_XLSX, payload)
    print(f"Wrote {config.OUT_XLSX}")
    return config.OUT_XLSX

if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run the unit test to verify it passes**

Run: `python3 -m pytest tests/test_urlmap_report.py -v`
Expected: PASS (1 passed)

- [ ] **Step 5: Run the FULL suite, then a fast offline end-to-end, then the live build**

```bash
cd /Users/aakash/Documents/Competator
python3 -m pytest -v                                   # all tasks green
python3 -c "from keyword_bible import build; build.main(run_live=False)"   # offline: parse+classify+excel, sitemaps still fetched
```
Expected offline run: prints a RECONCILIATION line whose buckets sum to total (assert passes), prints sitemap URL counts (1061 / 435), and writes `PolicyBazaar_UAE_Health_Travel_Keyword_Bible.xlsx`. Then run the live indexability build (slower — ~1,500 URL fetches):
```bash
python3 -c "from keyword_bible import build; build.main(run_live=True)"
```
Expected: same, with populated `Live:` columns.

- [ ] **Step 6: Manual verification & spot-check (evidence before claiming done)**

Open the workbook; confirm: 8 sheets present; Health/Travel matrices have all 8 competitor columns; Gaps sheets sorted by Opportunity descending; a blank competitor cell renders blank (not 0); GSC columns read `GSC not connected` (until wired). Spot-check 5 keywords per bucket against the raw SEMrush data. Record findings.

- [ ] **Step 7: Commit**

```bash
cd /Users/aakash/Documents/Competator
git add keyword_bible/build.py keyword_bible/urlmap_report.py tests/test_urlmap_report.py \
        "PolicyBazaar_UAE_Health_Travel_Keyword_Bible.xlsx"
git commit -m "feat: end-to-end build orchestrator with QA reconciliation + URL-map"
```

---

## Post-build follow-ups (not tasks; require your input)

- **GSC wiring (G4 actual status):** confirm the connected GSC property for `policybazaar.ae`; then replace `indexability.gsc_status` with the real URL Inspection adapter and re-run. Until then the workbook ships honestly labelled `GSC not connected`.
- **Seed-list & Review tuning:** review the README seed lists and the Review sheet; add/adjust terms and re-run `build.main()` (idempotent) to fold approved keywords into Health/Travel.

## Self-Review (completed by plan author)

- **Spec coverage:** G1 → Tasks 1,2,3,9 (ingest, classify, subtopics, assemble). G2 → Tasks 5,6 + assemble (sitemap fetch, suggest URL). G3 → Task 4 + assemble (tiers, opportunity, gap sheets). G4 → Task 7,10 (live check + GSC fallback, URL-map sheets). Output workbook (8 sheets) → Task 8. QA reconciliation & no-fabrication → Tasks 9,10 + README. Review tab → Tasks 9,10. All spec sections mapped.
- **Placeholder scan:** no TBD/TODO; every code/test step contains real content.
- **Type consistency:** `MATRIX_COLUMNS`/`GAP_COLUMNS`/`URLMAP_COLUMNS`/`REVIEW_COLUMNS` names match across `assemble.py`, `urlmap_report.py`, `excel.py`, and `build.py`; `gap_tier`/`opportunity`/`classify`/`suggest`/`check_url`/`gsc_status` signatures are consistent with their call sites in `assemble.build` and `build.main`.
