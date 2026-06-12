# РСВУ · Web of Science indicator verifier

Independently recompute and verify the **Web of Science science indicators** of
the Bulgarian University Ranking System (Рейтингова система на висшите училища в
България) using the **Web of Science Expanded API**, and compare them against the
published values in `data/Информация_ПН_ВУ_*.xlsx`.

It ships as a one-click **Streamlit web app** (deployable free) and an equivalent
**command-line** runner. The bundled data is for Medical University of Varna; the
matrix and values files can be swapped for any institution.

---

## Why the API (and not a spreadsheet export)

The Web of Science / InCites Excel exporter does **not** populate the
Core-Collection *Times Cited* column (it returns blanks or zeros), so an export
can only supply InCites citation counts — a broader citation universe that runs
roughly 20% higher than what РСВУ uses. The Expanded API returns the
Core-Collection times-cited per record, so the citation-based indicators are
computed on the correct basis.

---

## Indicators verified

Six indicators per professional field (ПН), matching columns L, N, P, R, T, U of
the values file:

| Col | Indicator (BG) | Meaning | Source |
|-----|----------------|---------|--------|
| L | Индекс на цитируемост на Хирш (WoS) | h-index | times-cited (WOS silo) |
| N | Среден брой цитирания на документ (WoS) | mean citations / document | times-cited (WOS silo) |
| P | Документи, цитирани поне веднъж (WoS) | documents cited ≥ 1 | times-cited (WOS silo) |
| R | Статии в научни списания (WoS) | journal articles | document type = `Article` |
| T | Доклади от научни конференции (WoS) | conference proceedings | document type = `Proceedings Paper` |
| U | Съвместни научни публикации (WoS) | co-authored with ≥ 1 other institution | ≥ 2 distinct institutions in addresses |

---

## How it works

A single master query fetches every record for the institution in the year
window:

```
OG=(Medical University Varna) AND PY=(2021-2025)
```

Each record is parsed once, then assigned to a professional field whenever any of
its WoS **Research Areas** maps to that ПН in `data/WoS_Area_pn_MUV.xlsx`. Research
areas are read from `category_info → subjects → ascatype="extended"`; a record can
carry several, so it can legitimately count toward several fields (e.g. a paper in
both *General & Internal Medicine* and *Public, Environmental & Occupational
Health* counts in Medicine, Public Health and Health Care). The five article- and
citation-level indicators are then recomputed per field and compared.

Key field choices:

- **Times-cited** — the `WOS` silo of
  `dynamic_data.citation_related.tc_list.silo_tc` (`coll_id="WOS"`) = Core
  Collection. h-index, mean citations and cited-once are derived from it.
- **Document type** — `static_data.summary.doctypes.doctype` drives R (`Article`)
  and T (`Proceedings Paper`).
- **Collaboration (U)** — distinct institutions across the record's addresses;
  ≥ 2 means at least one partner beyond the home institution. The API has no
  ready-made collaboration flag, so this is computed from the address data.

Area matching is done in punctuation-free normalized space, so it is unaffected
by spacing, commas or dashes — including research-area names that themselves
contain commas (`Public, Environmental & Occupational Health`,
`Dentistry, Oral Surgery & Medicine`).

---

## Install

```bash
pip install -r requirements.txt
```

Requirements: `streamlit`, `requests`, `openpyxl`, `xlrd`, `pandas`.

Your **WoS Expanded API key** is read from the environment (or Streamlit
secrets); it is never written to disk or committed. **Never paste it into a chat
or the code.**

```bash
export WOS_API_KEY="your-expanded-api-key"
```

---

## Step 1 — confirm the field paths (probe)

Expanded record JSON nests slightly differently across subscriptions, so confirm
the extractors before a full run:

```bash
python probe.py                 # 2 records: shows the extracted fields
python probe.py --raw           # also dumps the full raw JSON
python probe.py --ut WOS:000XXXXXXXXXXXX   # a specific record
python probe.py --n 3 --org 'OG=(Medical University Varna)' --from 2021 --to 2025
```

For each record it prints what the extractors pull: `times_cited(WOS)`,
`doctypes`, `research_areas(extended)`, `num_institutions`. If anything looks
wrong, adjust the isolated extractors in `core.py` — `extract_times_cited`,
`extract_doctypes`, `extract_research_area_keys`, `extract_num_institutions`.

---

## Step 2 — run the verification

### Web app

```bash
streamlit run app.py
```

Enter the API key (or set it in secrets), adjust the organization query / year
window / collaboration threshold, click **Run verification**. Results show a
green/red match table, an "indicators matching X / N" metric, a downloadable
`.xlsx` report, and a data-diagnostics panel (document-type counts, docs matching
no mapped area). A **Probe one record** button does Step 1 from the UI.

### Command line

```bash
python wos_verify.py --selftest         # offline logic checks, no API
python wos_verify.py --probe            # one record + extracted fields
python wos_verify.py --probe --raw      # ... plus raw JSON
python wos_verify.py                     # all fields -> wos_verification_report.xlsx
python wos_verify.py --pn 704            # a single field
python wos_verify.py --org 'OO=("Med Univ Varna" OR "Medical University of Varna")' \
                     --from 2021 --to 2025 --collab-min 2 --count-reviews
```

CLI flags: `--org`, `--from` / `--to`, `--collab-min`, `--count-reviews`,
`--pn`, `--info`, `--matrix`, `--out`, `--probe` / `--raw`, `--selftest`.

---

## Deploy free on Streamlit Community Cloud

1. Push this folder to a GitHub repo (public or private).
2. Go to **share.streamlit.io → Create app**, pick the repo and `app.py`.
3. Optionally, in **Advanced settings → Secrets**, paste:
   ```toml
   WOS_API_KEY = "your-expanded-api-key"
   ```
4. Deploy. You get a public `*.streamlit.app` URL; pushes auto-redeploy.

The matrix and values files are bundled, so the app works out of the box.

> ⚠ **Public-link caveat.** If you store the key in Secrets, every visitor
> consumes *your* WoS quota. For a shared link, leave Secrets empty so each
> viewer enters their own key, or restrict access with the Community Cloud viewer
> allow-list. `.streamlit/secrets.toml` is git-ignored and must never be
> committed (only `secrets.toml.example` is in the repo).

---

## Configuration reference (`core.Config`)

| Field | Default | Purpose |
|-------|---------|---------|
| `base_url` | `https://wos-api.clarivate.com/api/wos` | Expanded API endpoint |
| `database` | `WOS` | Core Collection |
| `edition` | `""` | restrict to a WoS edition (empty = all in your subscription) |
| `org_query` | `OG=(Medical University Varna)` | organization filter (`OG=` or `OO=`) |
| `py_from` / `py_to` | `2021` / `2025` | publication-year window |
| `article_doctypes` | `("Article",)` | document types counted as articles (R) |
| `proceedings_doctypes` | `("Proceedings Paper",)` | document types counted as proceedings (T) |
| `count_review_as_article` | `False` | include Reviews in R |
| `collab_min_institutions` | `2` | min distinct institutions for U |
| `page_size` | `100` | records per API page (Expanded max) |
| `sleep_between_calls` | `1.0` | pause between paged calls |
| `max_retries` | `4` | retry attempts on transient errors / 429 |

---

## Interpreting results

- **Document set** — `OG=` folds in indexed name variants. If totals look low,
  your institution may not be fully unified in WoS; switch to an
  `OO=("Variant 1" OR "Variant 2" …)` address search.
- **Collaboration (U)** — counted from address institutions, the metric most
  sensitive to WoS institution disambiguation; expect it to be the closest to,
  but not always exactly, the published value.
- **Matrix version** — the bundled matrix corresponds to the РСВУ 2019/2025
  correspondence table for MU-Varna's fields; small differences against another
  edition can shift area assignment.

---

## Files

```
app.py            Streamlit GUI (Expanded API)
core.py           verification logic + isolated, self-tested extractors
probe.py          confirm API field paths on your subscription
wos_verify.py     command-line runner
requirements.txt  dependencies
.streamlit/
   config.toml             theme
   secrets.toml.example    template (real secrets.toml is git-ignored)
data/
   WoS_Area_pn_MUV.xlsx           ПН ⇄ WoS Research-Area matrix (MU-Varna)
   Информация_ПН_ВУ_391003.xlsx   published values being verified
```

---

## License

MIT — see `LICENSE`.
