# Design Spec — PolicyBazaar UAE · Health & Travel Keyword Bible

- **Date:** 2026-06-30
- **Owner:** Aayush Gupta (SEO / Growth, PolicyBazaar UAE)
- **Status:** Draft for review
- **Scope of this spec:** Deliverable 1 only — the Excel "Keyword Bible" covering Goals G1–G4. Internal linking (G5) is explicitly deferred to a separate spec.

---

## 1. Background & context

We have 9 SEMrush "Organic Positions" exports, UAE database (`ae`), snapshot dated **2026-06-29** (exported 2026-06-30). One is PolicyBazaar's own site; eight are competitors. Every file has the same 17 columns: Keyword, Position, Previous position, Search Volume, Keyword Difficulty, CPC, URL, Traffic, Traffic (%), Traffic Cost, Competition, Number of Results, Trends, Timestamp, SERP Features by Keyword, Keyword Intents, Position Type.

**Row counts (verified by full ingest, not sampled):**

| Site | Role | Keywords | Est. traffic |
|---|---|---:|---:|
| policybazaar.ae | **Home (PB)** | 23,316 | 28,569 |
| mediclinic.ae | Competitor (hospital) | 35,825 | 500,465 |
| damaninsurance.ae | Competitor (health) | 4,933 | 73,799 |
| insuranceuae.com | Competitor (Orient) | 4,717 | 48,290 |
| sukoon.com | Competitor (multi-line) | 4,444 | 44,000 |
| insurancemarket.ae | Competitor (aggregator) | 10,863 | 38,015 |
| giggulf.ae | Competitor (multi-line) | 4,850 | 32,945 |
| shory.com | Competitor (car insurtech) | 8,621 | 32,592 |
| metlife.ae | Competitor (life) | 2,823 | 14,080 |

**Sitemaps (verified reachable, HTTP 200, flat urlsets):**

| Sitemap | PB URLs |
|---|---|
| https://www.policybazaar.ae/sitemap/sitemap-health-insurance.xml | 1,061 |
| https://www.policybazaar.ae/sitemap/sitemap-travel-insurance.xml | 435 |
| **Total** | **1,496** |

URL slugs carry topic suffixes (`-hiart` = health insurance article, `-hipvd` = health insurance product/provider detail, `-tipvd` = travel insurance provider detail, `-tic` = travel insurance content, etc.) used to auto-derive sub-topic clusters.

---

## 2. Goals (in scope)

- **G1 — Keyword mapping:** Extract and classify the full Health-Insurance and Travel-Insurance keyword universe across all 9 files.
- **G2 — URL mapping:** Map each keyword onto the correct PB URL from the two sitemaps (1,496 URLs).
- **G3 — Keyword gaps:** Identify where competitors rank and PB doesn't (or ranks poorly), with a suggested target URL.
- **G4 — Indexability:** Report indexability status of the PB Health + Travel URLs (live technical check + GSC actual index status).

**Success criterion:** A single, auditable Excel workbook that acts as the "bible" for Health + Travel keywords and mapping — comprehensive (no relevant keyword silently dropped), deterministic, and free of guessed/fabricated data.

## 2a. Non-goals / deferred

- **G5 — Internal linking** (Screaming Frog one-time suggestions + the ongoing internal-linking skill/tool). Separate spec, after Screaming Frog export is shared.
- Other verticals (car, life, home, pet, business, investment). Out of scope.
- Historical trend analysis beyond the single 2026-06-29 snapshot (the `Previous position` column is available and may be added as an optional column, but multi-period tracking is out of scope).

---

## 3. Inputs

1. The 9 SEMrush `.xlsx` files in `/Users/aakash/Documents/Competator/`.
2. The 2 PB sitemaps (fetched live; 1,496 URLs).
3. Google Search Console URL Inspection API for `policybazaar.ae` (for G4 actual index status). **Dependency:** GSC must be connected with access to the correct property. If unavailable, the GSC columns are filled with the literal value `GSC not connected` — never fabricated.

---

## 4. Methodology

All processing is **deterministic Python 3 + openpyxl**, scripted and re-runnable. No LLM is used in the classification or scoring path (the Hybrid method was chosen, not the LLM-assisted variant). Every transformation is auditable; intermediate results are saved as JSON.

### 4.1 Keyword universe construction

- Pool all rows from all 9 files.
- **Normalisation key** for de-duplication: keyword lowercased, trimmed, internal whitespace collapsed. Arabic keywords are retained as-is (kept, not dropped).
- **Per-site de-dupe:** one record per (site, normalised-keyword). Where a site has multiple rows for the same keyword (different URLs), keep the **best (lowest) position**; PB's ranking URL = the URL of PB's best-position row.
- The master keyword list = the union of normalised keywords across all sites that are classified Health or Travel (see 4.2).

### 4.2 Classification (Hybrid — URL-path + word list + Review)

Applied per unique keyword, in this deterministic order:

1. **URL-path signal** (strongest): if any site's ranking URL for the keyword contains a Health marker (`/health-insurance/`, `-hiart`, `-hipvd`, `/medical`, hospital paths) → **Health**; a Travel marker (`/travel-insurance/`, `-tipvd`, `-tic`, `/travel`) → **Travel**.
2. **Keyword word-list signal:** match against curated Health and Travel term lists (see 4.3).
3. **Conflict** (both Health and Travel signals present) → **Review** (flagged with both candidate tags).
4. **Neither signal**, but matches a hard-exclude term (car/banking/etc.) → **Out-of-scope** (dropped from bible, retained in an audit count).
5. **Neither signal and not hard-excluded** → **Review** (so nothing relevant is silently lost).

Every input keyword lands in exactly one bucket: Health, Travel, Out-of-scope, or Review. Bucket counts must reconcile to the total (QA check, §7).

### 4.3 Seed word lists (user-approvable, editable)

These are seeds; the user reviews and edits them, and reviews the Review tab. They are stored in the README sheet for transparency.

- **Health seeds (EN):** health insurance, medical insurance, health, medical, healthcare, thiqa, daman, mediclinic, clinic, hospital, dha, doh, haad, doctor, treatment, medicine, pharmacy, dental, maternity, pregnancy, ivf, opd/ip, chronic, illness, disease, copay, network, reimbursement, enaya, saada, basmah, isahd, mednet, nextcare, family health, individual/group health, health card, checkup, vaccination, + insurer health-plan brand names.
- **Health seeds (AR):** تأمين صحي, ضمان, صحي, مستشفى, علاج, طبي.
- **Travel seeds (EN):** travel, visa, schengen, trip, passport, tourist, holiday, vacation, abroad, overseas, baggage, luggage, immigration, green list, quarantine, single/multi/annual trip, travel medical, emergency abroad, worldwide/international travel, + country-name + visa patterns.
- **Travel seeds (AR):** تأمين سفر, سفر, تأشيرة, فيزا.
- **Hard-exclude seeds:** car, motor, vehicle, mulkiya, rta, traffic fine, salik, tasjeel, driving license, bike, bank, atm, balance check, credit card, loan, salary, ratibi, postal code, life insurance, home insurance, pet insurance, investment, trading, savings. (A hard-excluded keyword can still be rescued into Health/Travel by a strong URL-path signal — step 1 wins over step 4.)

### 4.4 Vertical scope (Core + adjacent, intent-tagged)

Included: (a) commercial insurance terms and (b) the adjacent content/article topics PB and competitors use to attract the audience — Health: hospitals, clinics, DHA/DOH, conditions, procedures; Travel: visas, passports, country guides, schengen. Each keyword is tagged by **Intent** (informational / commercial / transactional / navigational) so the user can filter. **Competitor brand-navigational junk is excluded** (e.g. `mediclinic parkview hospital`, `daman login`) — a competitor-brand + navigational-intent keyword that PB would never target is dropped to Out-of-scope (PB's own brand terms are retained).

### 4.5 Sub-topic clustering

Each keyword gets a **Sub-topic** label derived deterministically from (a) the PB URL slug suffix/tokens where PB ranks, and (b) keyword token patterns. Examples: `core-health-insurance`, `hospitals`, `gov-plan (thiqa/enaya/saada)`, `family-health`, `dental`, `claims-reimbursement`, `provider-brand`; `core-travel-insurance`, `schengen-visa`, `country-guide`, `passport-visa`, `covid-travel`. Unmatched → `uncategorised`.

### 4.6 Competitive matrix

One row per unique in-scope keyword. Columns hold each site's best position (blank where the site does not appear in the dataset for that keyword).

> **Data-integrity note (documented in README):** Each SEMrush export only contains keywords where *that* site ranks in the top 100 as of the snapshot. Therefore a blank cell means "not present in that site's top-100 export on 2026-06-29" — **not** a guaranteed absence from Google. This is stated explicitly; blanks are never presented as "definitely does not rank."

### 4.7 Gap logic (G3) — tiered + volume-weighted

Per keyword, comparing PB against the best competitor position:

- **Missing:** PB has no ranking in the dataset; ≥1 competitor ranks.
- **Weak:** PB ranks but position > 20, while the best competitor is ≤ 20.
- **Near:** PB ranks ≤ 20 but at least one competitor outranks PB.
- **Winning:** PB is the best-positioned site (PB ≤ every ranking competitor), or PB ranks and no competitor does.

**Opportunity score** (transparent, documented formula):

```
Opportunity = round( Search_Volume × severity_weight × intent_multiplier )
severity_weight   = { Missing:1.0, Weak:0.7, Near:0.4, Winning:0.0 }
intent_multiplier = { transactional:1.2, commercial:1.15, informational:1.0, navigational:0.6 }
```

Gaps sheets are sorted by Opportunity descending.

### 4.8 URL mapping (G2) & suggested URL

- If PB already ranks: **Suggested URL = PB's ranking URL**; validate it exists in the sitemap (flag `not-in-sitemap` if not).
- If Missing (or PB URL not in sitemap): find the best-matching sitemap URL by deterministic **token-overlap score** between keyword tokens and URL-slug tokens (plus sub-topic match). If the best match scores below a set threshold → **`new page needed`**. Low-confidence matches are flagged `low-confidence — verify`. No URL is invented; suggestions come only from the 1,496 real sitemap URLs or the explicit "new page" flag.

### 4.9 Indexability (G4)

Checked for the PB Health + Travel URLs (the 1,496 sitemap URLs, plus any PB ranking URL found in the data but missing from the sitemap).

- **Live technical check (no auth):** HTTP status, final URL after redirects, robots.txt allow/deny, meta-robots / X-Robots-Tag `noindex`, canonical URL + self-canonical (Y/N), and a derived **Indexable verdict (Y/N + reason)**.
- **GSC actual status (needs connection):** coverage state (e.g. "Submitted and indexed", "Crawled – currently not indexed", "Discovered – currently not indexed", "URL is unknown to Google"), last crawl date, Google-selected canonical. If GSC is unavailable → `GSC not connected` in every GSC cell (honest, not blank/faked).

---

## 5. Output — `PolicyBazaar UAE · Health & Travel Keyword Bible.xlsx`

Single workbook, 8 sheets:

| # | Sheet | Contents |
|---|---|---|
| 1 | **README / Methodology** | Snapshot date, competitor list, classification rules, seed word lists, gap definitions, opportunity formula, colour legend, and the top-100 data-integrity caveat |
| 2 | **Health — Master Matrix** | One row per unique Health keyword |
| 3 | **Travel — Master Matrix** | One row per unique Travel keyword |
| 4 | **Health — Gaps (prioritised)** | Missing/Weak/Near rows, sorted by Opportunity, with suggested URL |
| 5 | **Travel — Gaps (prioritised)** | Same, Travel |
| 6 | **Health — URL Map** | One row per PB Health URL (see columns below) |
| 7 | **Travel — URL Map** | Same, Travel |
| 8 | **Review** | Ambiguous keywords for user sign-off, with candidate tags and reason |

### 5.1 Master Matrix columns (sheets 2–3)

Keyword · Vertical · Sub-topic · Intent · Search Volume · Keyword Difficulty · CPC · **PB Position** · **PB Ranking URL** · **PB Suggested/Mapped URL** · Daman · Sukoon · Orient · GIG · MetLife · InsuranceMarket · Shory · Mediclinic (positions) · Best Competitor · Best Competitor Position · **Gap Status** · **Opportunity Score** · SERP Features · **AI Overview? (Y/N)** · Suggested Action.

- All 8 competitors get a position column (blank = not in that site's top-100 export).
- `AI Overview?` is derived from the SERP-features field (AI Overview presence is a GEO signal worth surfacing).

### 5.2 URL Map columns (sheets 6–7)

PB URL · Sub-topic · # Keywords mapped · Sum Traffic · Top keyword (by traffic) · Best position on page · In sitemap? · **Live: HTTP status** · **Live: robots.txt allowed** · **Live: noindex?** · **Live: canonical / self-canonical** · **Live: Indexable verdict + reason** · **GSC: coverage state** · **GSC: last crawl** · **GSC: Google-selected canonical** · Notes (orphan / thin / duplicate-target).

---

## 6. Build sequencing

1. **Now (all data in hand):** Sheets 1, 2, 3, 8 in full + gap tiers, opportunity scores, competitive matrix, SERP/AI-Overview flags. G2 URL mapping using the fetched sitemaps → fills Suggested URL and sheets 6–7 URL side. Live indexability check (G4 technical half).
2. **GSC connected:** fill GSC columns in sheets 6–7. If not connected, ship with `GSC not connected` and backfill later.

Nothing else is blocked.

---

## 7. Data-integrity & QA checks (mandatory)

- **Reconciliation:** total input rows → (Health + Travel + Out-of-scope + Review) bucket counts must sum with zero loss; report the counts.
- **No fabrication:** blank position = "absent from that site's top-100 export" (documented). GSC failure = explicit `GSC not connected`. Sitemap/GSC/HTTP fetch failure → stop and report per the user's tool-failure policy; do not infer.
- **Suggested URLs** come only from the real sitemap set or the explicit `new page needed` flag.
- **Spot-check:** manual review of a keyword sample per bucket before shipping.
- **Reproducibility:** the Python build script + intermediate JSON are saved so the workbook can be regenerated identically.

---

## 8. Open items / assumptions

- **GSC access:** confirm the connected property for `policybazaar.ae` (domain vs URL-prefix) before the G4 GSC step.
- **Seed word lists & Review tab:** user to review/approve (§4.3, sheet 8).
- **Language:** both English and Arabic keywords retained; `ar-ae/` PB URLs included in mapping.
- **Snapshot:** single point in time (2026-06-29); no multi-period trending in this deliverable.

---

## 9. Out of scope (restated)

G5 internal linking — both the one-time Screaming Frog suggestions and the ongoing internal-linking skill/tool — will be brainstormed and specced separately once the Screaming Frog export is shared.
