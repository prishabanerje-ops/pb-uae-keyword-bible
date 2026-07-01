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
