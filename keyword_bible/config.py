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
