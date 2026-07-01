"""
build.py — End-to-end pipeline orchestrator for the Health & Travel Keyword Bible.

Usage:
    python3 -c "from keyword_bible import build; build.main(run_live=False)"
    python3 -c "from keyword_bible import build; build.main(run_live=True)"
    python3 -m keyword_bible.build
"""

import json
import os
from . import config, ingest, sitemaps, assemble, excel, indexability, urlmap_report


def _gap_rows(matrix_rows):
    """Return sorted gap rows for the Gap sheet (Missing / Weak / Near only)."""
    keep = {"Missing", "Weak", "Near"}
    rows = [r for r in matrix_rows if r.get("Gap Status") in keep]
    rows.sort(key=lambda r: r.get("Opportunity Score", 0), reverse=True)
    return [{c: r.get(c, "") for c in assemble.GAP_COLUMNS} for r in rows]


def main(run_live: bool = True) -> str:
    """Run the full pipeline and write the output workbook.

    Args:
        run_live: if True, fetch live indexability for every URL in the sitemaps
                  (~1 500 HTTP requests).  If False, skip live checks so the
                  pipeline can be validated offline in seconds.

    Returns:
        Absolute path to the written workbook.
    """
    os.makedirs(config.INTERIM_DIR, exist_ok=True)

    print("Loading SEMrush files...")
    rbs = ingest.load_all()

    print("Fetching sitemaps...")
    sm = sitemaps.load_all()
    print(f"  health sitemap: {len(sm['health'])} urls | travel sitemap: {len(sm['travel'])} urls")

    data = assemble.build(rbs, sm)
    recon = data["recon"]
    print(
        f"RECONCILIATION: total={recon['total']} health={recon['health']} "
        f"travel={recon['travel']} out={recon['out']} review={recon['review']}"
    )
    assert recon["health"] + recon["travel"] + recon["out"] + recon["review"] == recon["total"], \
        "Reconciliation failed — bucket counts do not sum to total"

    # Indexability functions — real or stub
    if run_live:
        index_fn = indexability.check_url
    else:
        def index_fn(u):  # noqa: E306
            return {
                "http_status": "SKIPPED", "final_url": u, "robots_allowed": "",
                "noindex": "", "canonical": "", "self_canonical": "",
                "indexable": "", "reason": "live check skipped",
            }

    gsc_fn = indexability.gsc_status

    print("Building URL-Map (live indexability)..." if run_live else "Building URL-Map (skipped live)...")
    health_urlmap = urlmap_report.build_rows(sm["health"], "Health", data["health"], index_fn, gsc_fn)
    travel_urlmap = urlmap_report.build_rows(sm["travel"], "Travel", data["travel"], index_fn, gsc_fn)

    readme = [
        "PolicyBazaar UAE — Health & Travel Keyword Bible",
        f"Snapshot date: {config.SNAPSHOT_DATE}",
        (
            f"Reconciliation: total={recon['total']} health={recon['health']} "
            f"travel={recon['travel']} out={recon['out']} review={recon['review']}"
        ),
        "Gap tiers: Missing / Weak / Near / Winning. Opportunity = Volume x severity x intent multiplier.",
        (
            "IMPORTANT: a blank competitor position = 'absent from that site's top-100 SEMrush export "
            "on the snapshot date', NOT a guaranteed non-ranking."
        ),
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

    # Dump interim JSON (matrix + review rows; no recon key)
    interim_path = os.path.join(config.INTERIM_DIR, "assembled.json")
    with open(interim_path, "w", encoding="utf-8") as f:
        json.dump({k: v for k, v in data.items() if k != "recon"}, f, ensure_ascii=False)

    excel.write_workbook(config.OUT_XLSX, payload)
    print(f"Wrote {config.OUT_XLSX}")
    return config.OUT_XLSX


if __name__ == "__main__":
    main()
