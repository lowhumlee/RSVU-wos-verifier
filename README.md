# РСВУ · Web of Science indicator verifier (Expanded API)

Recompute and verify the **Web of Science science indicators** of the Bulgarian
University Ranking System (Рейтингова система на висшите училища в България)
against the **Web of Science Expanded API**, and compare them with the published
values in `data/Информация_ПН_ВУ_*.xlsx`.

The API is used (instead of a spreadsheet export) because it returns the
**Core-Collection times-cited** per record — the count РСВУ actually used — while
the InCites/WoS exporter left that column empty.

Five indicators per professional field (ПН):

| Col | Indicator (BG) | Meaning | Source field |
|-----|----------------|---------|--------------|
| L | Индекс на цитируемост на Хирш (WoS) | h-index | times-cited (WOS silo) |
| N | Среден брой цитирания на документ (WoS) | mean citations/doc | times-cited (WOS silo) |
| P | Документи, цитирани поне веднъж (WoS) | docs cited ≥1 | times-cited (WOS silo) |
| R | Статии в научни списания (WoS) | journal articles | document type = Article |
| U | Съвместни научни публикации (WoS) | co-authored w/ ≥1 other institution | ≥2 institutions in addresses |

## How it works

One master query fetches every university record in the window:

```
OG=(Medical University Varna) AND PY=(2021-2025)
```

Each record is parsed and assigned to a professional field when its WoS
**Research Area** (`category_info → subjects → ascatype="extended"`) maps to that
ПН in `data/WoS_Area_pn_MUV.xlsx`. The five indicators are then recomputed per
field and compared.

Times-cited is taken from the **WOS** silo
(`dynamic_data.citation_related.tc_list.silo_tc`, `coll_id="WOS"`) = Core
Collection. Collaboration counts documents with ≥2 distinct institutions in their
addresses.

## Confirm the field paths first (probe)

Record JSON nesting varies slightly by subscription. Confirm before trusting
numbers — never paste your key into a chat; read it from the environment:

```bash
pip install -r requirements.txt
export WOS_API_KEY="your-key"
python probe.py                 # 2 records: extracted fields side-by-side
python probe.py --raw           # also dump full raw JSON
```

If any extracted field is wrong, adjust the isolated extractors in `core.py`:
`extract_times_cited`, `extract_research_area_keys`, `extract_num_institutions`,
`extract_doctypes`.

## Run

```bash
streamlit run app.py                 # GUI
# or:
python wos_verify.py --selftest      # offline checks
python wos_verify.py --probe         # one record, extracted fields
python wos_verify.py                  # all fields -> report.xlsx
python wos_verify.py --pn 704 --collab-min 2
```

## Deploy free on Streamlit Community Cloud

1. Push to GitHub. 2. **share.streamlit.io → Create app →** pick repo + `app.py`.
3. In **Advanced settings → Secrets** paste `WOS_API_KEY = "…"` (or leave empty so
each viewer enters their own key). `.streamlit/secrets.toml` is git-ignored.

⚠ If you store the key in Secrets, every visitor consumes your WoS quota — for a
public link, prefer bring-your-own-key or the viewer allow-list.

## Files

```
app.py            Streamlit GUI (Expanded API)
core.py           logic + isolated extractors (self-tested)
probe.py          confirm API field paths on your subscription
wos_verify.py     CLI runner
data/
   WoS_Area_pn_MUV.xlsx           ПН ⇄ WoS Research-Area matrix (MU-Varna)
   Информация_ПН_ВУ_391003.xlsx   values being verified
```

## License

MIT — see `LICENSE`.
