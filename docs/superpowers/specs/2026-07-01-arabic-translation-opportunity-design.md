# Design Spec — PolicyBazaar UAE · Arabic Translation Opportunity (Health & Travel)

- **Date:** 2026-07-01
- **Owner:** Aayush Gupta (SEO / Growth, PolicyBazaar UAE)
- **Status:** Draft for review
- **Approach:** A — Prioritized Translation Backlog (approved 2026-07-01)
- **Scope of this spec:** Identify and rank which existing English Health & Travel pages to translate into Arabic, backed only by measured data, to grow Arabic organic visibility and E-E-A-T. Includes a health-check of the 32 existing Arabic pages. Net-new Arabic-only demand is a light, secondary tab.

---

## 1. Background & context

PolicyBazaar UAE publishes its Health and Travel content almost entirely in English. Arabic coverage is negligible, while Arabic search demand for these topics is large and — where PB has built Arabic pages — already converts. The goal is a single, auditable "which pages to translate first" workbook, scored by real data, mirroring the existing Keyword Bible deliverable.

### 1.1 Verified facts (measured, not inferred)

**Sitemaps** (fetched live 2026-07-01, HTTP 200):

| Sitemap | Total URLs | Arabic (`ar-ae`) URLs |
|---|---:|---:|
| sitemap-health-insurance.xml | 1,061 | 31 |
| sitemap-travel-insurance.xml | 435 | 1 |
| **Total** | **1,496** | **32 (~2%)** |

No dedicated Arabic sitemap exists; Arabic URLs live under the `/ar-ae/` path inside each vertical sitemap. Travel is effectively 0% translated.

**Arabic search demand in the SEMrush exports** (UAE db `ae`, snapshot 2026-06-29; Arabic-script keywords only):

| Site | Arabic keywords | Arabic search volume |
|---|---:|---:|
| **policybazaar.ae (PB)** | **2,182** | **694,500** |
| mediclinic.ae | 7,038 | 3,522,150 |
| damaninsurance.ae | 1,193 | 724,010 |
| shory.com | 1,637 | 673,800 |
| insurancemarket.ae | 611 | 215,600 |
| sukoon.com | 339 | 112,560 |
| insuranceuae.com (Orient) | 442 | 97,750 |
| giggulf.ae | 439 | 89,080 |
| metlife.ae | 70 | 14,770 |

**Headline signal:** PB already ranks for 2,182 Arabic queries (~694K volume) with only 32 Arabic pages — Google is serving PB's **English** pages to Arabic searchers. An Arabic page closes that relevance/E-E-A-T gap.

**GSC proof points** (live API, 90-day sample, `sc-domain:policybazaar.ae`): where PB has Arabic Health pages, they rank top-2 — e.g. `تأمين صحي لكبار السن في الإمارات` → `/ar-ae/senior-citizen-health-insurance-hic/` at **position 1.6**; `أفضل شركات التأمين الصحي في دبي` → the `/ar-ae/…/top-10-health-insurance…` article at **position 1.7**. PB's out-of-scope `/ar-ae/car-insurance/` pages pull the largest Arabic click volume (pos 2–11), demonstrating the lever works when Arabic pages exist.

### 1.2 Relationship to the existing Keyword Bible

The Keyword Bible pipeline (`keyword_bible/` modules, `gap-analysis` skill) is reused for ingest, classification, sitemap parsing, position matrix, and indexability. **However, its `assembled.json` dropped every Arabic keyword** (classification matched English seeds only → 0 Arabic rows in all buckets). Therefore Arabic keywords are **re-derived from the raw SEMrush files**; the Bible's `assembled.json` is not reused as an Arabic source.

---

## 2. Goals (in scope)

- **A1 — Arabic keyword universe:** extract and classify (Health/Travel) all Arabic-script keywords across the 9 SEMrush files, retaining volume, difficulty, per-site positions, and ranking URLs.
- **A2 — Arabic demand → page mapping:** pull real Arabic query × page data from GSC (country = United Arab Emirates) to find which English pages already receive Arabic-query impressions.
- **A3 — Translation priority:** rank existing English Health & Travel URLs by a transparent, measured `TranslateScore`, bucketed into P1/P2/P3.
- **A4 — Existing Arabic health-check:** audit the 32 live Arabic URLs for indexation, hreflang reciprocity, canonical correctness, and orphan status.
- **A5 — Arabic-only gaps (light):** surface Arabic queries with no good English *or* Arabic page as net-new candidates.

**Success criterion:** a single auditable Excel workbook that tells PB, in priority order, which Health & Travel pages to translate first — comprehensive, deterministic, and free of guessed/fabricated data.

## 2a. Non-goals / deferred

- Full Arabic content strategy (net-new page architecture, phased roadmap, effort estimates) — Approach B, a possible fast-follow after this ranked list is reviewed.
- Verticals other than Health and Travel.
- Revenue/leads weighting in the score — **explicitly excluded per owner decision** (2026-07-01). The `query_result_*.csv` leads exports are not used in `TranslateScore`.
- Producing the Arabic translations themselves; native-vs-machine translation quality is a flagged recommendation, not a build output.
- Multi-period trending; single snapshot (SEMrush 2026-06-29) plus a GSC trailing window.

---

## 3. Inputs

1. The 9 SEMrush `.xlsx` organic-position exports in `/Users/aakash/Documents/Competator/` (snapshot 2026-06-29).
2. The 2 PB sitemaps (fetched live; 1,496 URLs, 32 Arabic).
3. **GSC Search Analytics API** — Tier 2 OAuth, property `sc-domain:policybazaar.ae`, scopes include `webmasters`. Used for Arabic query × page × country demand and for English-page current performance.
4. **`Clicks and Impression.csv`** (`~/Downloads/`) — page × month clicks/impressions, all pages; used for English-page current performance where GSC page rows are unavailable.
5. **URL Inspection API** (Tier 1, same OAuth) — index status + Google-selected canonical for the 32 Arabic URLs.
6. Live HTTP fetch of the 32 Arabic URLs — hreflang, canonical, meta-robots.

**Dependency note:** GSC is confirmed connected (Tier 2, verified 2026-07-01). If any API/fetch fails, the run **stops and reports** per the owner's tool-failure policy; affected cells are labelled (e.g. `GSC not available`), never fabricated.

---

## 4. Methodology (deterministic Python 3; no LLM in the scoring path)

### 4.1 Arabic keyword extraction & classification

- Pool all rows from all 9 files. Detect Arabic-script keywords with regex covering the Arabic Unicode block (`U+0600–U+06FF`, extendable to `U+0750–U+077F`, `U+08A0–U+08FF`). The exact character class is documented in the README sheet.
- Classify each Arabic keyword Health / Travel / Out-of-scope / Review using the existing hybrid order: (1) URL-path marker on any site's ranking URL, (2) Arabic + English seed word lists, (3) conflict → Review, (4) hard-exclude (car/motor/bank/etc.) → Out-of-scope, (5) neither → Review. Arabic seeds are extended beyond the Bible's minimal set (e.g. `تأمين صحي`, `تأمين سفر`, `تأشيرة`, `فيزا`, `شنغن`, `مستشفى`, `علاج`, `صحي`, `طبي`, plus emirate and provider terms); the full list lives in the README sheet and is owner-editable.
- Per-site de-dupe keeps the best (lowest) position; PB's ranking URL = URL of PB's best-position Arabic row.

### 4.2 GSC Arabic demand → page mapping (A2)

- Extend `keyword_bible`/`seo-google` `gsc_query.py` to: filter **country = `are`**, request dimensions `query` + `page`, paginate beyond 1,000 rows (GSC `startRow`), over a trailing window (target 12 months; ≤16 per API limit).
- Flag Arabic-script queries; aggregate impressions / clicks / average position **per landing page**.
- A high volume of Arabic impressions landing on an **English** URL is the strongest translate-me signal (Google already ranks that page for Arabic intent).
- **Documented caveat:** GSC anonymizes low-volume queries, so some Arabic demand is hidden. This is stated in README; hidden demand is never estimated or back-filled.

### 4.3 English URL universe & sub-topics

- The 1,496 sitemap URLs, split into 32 existing Arabic (`/ar-ae/…`) and ~1,464 English-only. Sub-topic labels are derived from slug suffixes/tokens (`-hiart`, `-hipvd`, `-hic`, `-tipvd`, `-tic`, etc.), reusing the Bible's logic.
- Each English candidate URL is matched to (a) its mapped Arabic keywords (token/sub-topic overlap) and (b) its GSC Arabic impressions.

### 4.4 Scoring (A3) — demand + performance only

Per English Health/Travel URL:

```
TranslateScore = round( ArabicDemand × Performance × CompetitorGap × TwinFactor )

ArabicDemand   = Σ Arabic SEMrush search volume of the URL's mapped keywords
                 + GSC Arabic impressions landing on the URL
Performance    = the URL's current GSC impressions (proven live demand; falls back to
                 Clicks-and-Impression.csv impressions when a GSC page row is unavailable)
CompetitorGap  = 1 + (share of the 8 competitors that rank in Arabic for the URL's sub-topic)
TwinFactor     = 1.0 if no /ar-ae/ twin exists
                 0.2 if an Arabic twin already exists (deprioritise re-translation; keep visible)
```

- Normalisation, exact weights, and any log-scaling are fixed constants documented in the README sheet.
- Rows sorted by `TranslateScore` descending and bucketed: **P1** (top tier), **P2** (mid), **P3** (long tail). Tier cutoffs are documented, not eyeballed.

### 4.5 Existing Arabic health-check (A4)

For each of the 32 Arabic URLs:
- **URL Inspection API:** coverage/index state, Google-selected canonical, last crawl.
- **Live fetch:** HTTP status, self-canonical (Y/N), meta-robots/X-Robots `noindex`, and **reciprocal hreflang** — does the Arabic page point `hreflang="ar-ae"`↔`x-default`/`en` to its English twin, and does the English twin point back? Non-reciprocal or missing hreflang is a common E-E-A-T/indexation defect.
- **Twin mapping:** derive the expected English source URL (drop `/ar-ae/`), confirm it exists in the sitemap, flag orphans (Arabic page with no internal links / not in sitemap).
- Output a per-URL **verdict + concrete fix**.

### 4.6 Arabic-only gaps (A5, light)

Arabic queries (from SEMrush + GSC) whose best-matching PB page is weak/absent in **both** English and Arabic → net-new Arabic page candidates, with suggested `/ar-ae/` slug and the demand figure. Kept intentionally light (one tab, no scoring model beyond demand).

---

## 5. Output — `PolicyBazaar_UAE_Arabic_Translation_Opportunity.xlsx`

Single workbook, 6 sheets:

| # | Sheet | Contents |
|---|---|---|
| 1 | **README / Methodology** | Sources, snapshot & GSC window dates, Arabic-detection rule, seed lists, scoring formula + weights + tier cutoffs, GSC-anonymization caveat, colour legend, data-integrity notes |
| 2 | **Health — Translation Priority** | Ranked English Health URLs (columns in §5.1) |
| 3 | **Travel — Translation Priority** | Same, Travel |
| 4 | **Existing Arabic Pages — Health Check** | 32 Arabic URLs (columns in §5.2) |
| 5 | **Arabic Demand — Keyword Universe** | Arabic keywords by sub-topic; PB vs each competitor position; volume; intent |
| 6 | **Arabic-Only Gaps** | Net-new Arabic candidates: query, volume, suggested `/ar-ae/` slug, why no current page |

### 5.1 Translation Priority columns (sheets 2–3)

EN URL · Sub-topic · EN impressions (GSC) · EN clicks (GSC) · # Arabic keywords mapped · Arabic search volume (SEMrush) · Arabic impressions on page (GSC) · Best Arabic query + its position · AR twin exists? (Y/N + URL) · # Competitors ranking in Arabic · **TranslateScore** · **Priority tier (P1/P2/P3)** · Suggested AR URL (`/ar-ae/` + EN slug mirror) · Notes.

### 5.2 Health-check columns (sheet 4)

AR URL · Mapped EN source URL · EN source in sitemap? · Live HTTP status · Self-canonical? · `noindex`? · **hreflang reciprocal? (Y/N + detail)** · GSC coverage state · Google-selected canonical · Arabic impressions / clicks / avg position (GSC) · Orphan? · **Verdict + fix**.

---

## 6. Build sequencing

1. **Phase 1 (all local data in hand):** sheets 1, 2, 3, 5 + scoring from SEMrush + sitemaps + `Clicks and Impression.csv`. Arabic universe, classification, competitive matrix, sub-topics, provisional `TranslateScore` (GSC term = 0 until Phase 2).
2. **Phase 2 (GSC):** paginated Arabic query × page pull (country = are, 12-mo) → fills GSC columns, finalises `TranslateScore`, produces sheet 6.
3. **Phase 3 (indexability):** URL Inspection + live hreflang for the 32 Arabic URLs → sheet 4.

Nothing is blocked; GSC is confirmed available.

---

## 7. Data-integrity & QA (mandatory — owner policy)

- **Reconciliation:** total Arabic input rows → (Health + Travel + Out-of-scope + Review) must sum with zero loss; report counts.
- **No fabrication:** blank competitor position = "absent from that site's top-100 export" (documented). Hidden GSC queries are not estimated. Any API/HTTP/sitemap failure → **stop and report** what failed and what data is missing; do not infer.
- **Suggested AR URLs** are deterministic `/ar-ae/` mirrors of real English slugs — no invented pages or content.
- **Spot-check:** manual review of a sample per bucket and of the top-20 P1 rows before shipping.
- **Reproducibility:** Python build script + intermediate JSON saved so the workbook regenerates identically.

---

## 8. Testing

Unit tests mirror the existing `tests/` suite:
- Arabic-script detection (positive/negative, mixed-script strings).
- Classification of Arabic keywords into Health/Travel/Out/Review.
- `TranslateScore` computation and P1/P2/P3 bucketing on fixture rows.
- hreflang reciprocity parsing (reciprocal, one-way, missing).
- GSC-response parsing incl. pagination and country filter (mocked).

---

## 9. Open items / assumptions

- **GSC anonymization:** some Arabic queries are hidden by Google; documented, never estimated.
- **`/ar-ae/` pattern:** confirmed as PB's Arabic path from the sitemaps; suggested AR URLs mirror it.
- **GSC window:** target 12 months (≤16 per API); exact window recorded in README.
- **Translation quality:** native vs machine translation is a flagged recommendation, out of build scope.
- **Seed lists:** Arabic + English seeds are owner-editable in the README sheet before final scoring.

---

## 10. Reuse & new components

- **Reuse:** `keyword_bible/ingest.py`, `classify.py`, `sitemaps.py`, `indexability.py`, `excel.py`, `subtopics.py`; the `gap-analysis` skill engine.
- **New modules:** `arabic_extract` (Arabic-script detection + Arabic seed classification), `gsc_arabic` (country-filtered, paginated query×page pull + Arabic flagging), `translate_score` (the §4.4 model + tiering), `hreflang_check` (§4.5 reciprocity + twin mapping), and `arabic_excel` (the 6-sheet writer).
