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
