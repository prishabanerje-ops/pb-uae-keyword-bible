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
