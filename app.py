# -*- coding: utf-8 -*-
"""
app.py — Streamlit GUI to verify РСВУ Web of Science indicators via the
Web of Science Expanded API (Core-Collection times-cited). No upload of
citations needed; the API provides them.

Run locally:  streamlit run app.py
Deploy free:  push to GitHub -> share.streamlit.io -> Create app.

API key precedence: st.secrets["WOS_API_KEY"]  ->  sidebar password field.
The key is never written to disk or committed.
"""
import json
import os
from collections import Counter

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="РСВУ · WoS API verifier", page_icon="🔬", layout="wide")

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
DEFAULT_INFO = os.path.join(DATA_DIR, "Информация_ПН_ВУ_391003.xlsx")
DEFAULT_MATRIX = os.path.join(DATA_DIR, "WoS_Area_pn_MUV.xlsx")

st.title("🔬 РСВУ · Web of Science indicator verifier")
st.caption(
    "Recomputes the WoS science indicators of the Bulgarian University Ranking "
    "System from the Web of Science Expanded API (Core-Collection times-cited) and "
    "compares them with the published values."
)

with st.sidebar:
    st.header("Settings")
    try:
        secret_key = st.secrets.get("WOS_API_KEY", "")
    except Exception:
        secret_key = ""
    if secret_key:
        st.success("API key loaded from secrets.")
        api_key = secret_key
    else:
        api_key = st.text_input("WoS Expanded API key", value=os.environ.get("WOS_API_KEY", ""),
                                type="password", help="Used only this session; never stored.")
    st.divider()
    org_query = st.text_input("Organization query", value="OG=(Medical University Varna)",
                              help="OG= folds in indexed name variants. If WoS hasn't unified "
                                   "all spellings, use OO=(\"Variant 1\" OR \"Variant 2\").")
    c1, c2 = st.columns(2)
    py_from = c1.number_input("Year from", 1990, 2100, 2021, step=1)
    py_to = c2.number_input("Year to", 1990, 2100, 2025, step=1)
    collab_min = st.number_input("Collaboration: min distinct institutions", 2, 10, 2, step=1,
                                 help="U counts documents with at least this many institutions.")
    count_reviews = st.checkbox("Count Reviews as articles (R)", value=False)
    with st.expander("Override bundled matrix / values"):
        up_matrix = st.file_uploader("ПН ⇄ WoS area matrix (.xlsx/.xls)", type=["xlsx", "xls"])
        up_info = st.file_uploader("Информация…xlsx", type=["xlsx"])

cfg = core.Config(org_query=org_query, py_from=int(py_from), py_to=int(py_to),
                  collab_min_institutions=int(collab_min),
                  count_review_as_article=bool(count_reviews))

try:
    info = core.load_info(up_info if up_info else DEFAULT_INFO)
    mapping, keysets = core.load_area_matrix(up_matrix if up_matrix else DEFAULT_MATRIX)
except Exception as e:
    st.error("Could not read matrix / values: %s" % e); st.stop()

labels = {r["pn_code"]: "ПН%s · %s" % (r["pn_code"], r["pn_name"]) for r in info}
chosen = st.multiselect("Professional fields", list(labels.keys()),
                        default=list(labels.keys()), format_func=lambda c: labels[c])
info_sel = [r for r in info if r["pn_code"] in chosen]

with st.expander("Preview query & area mapping"):
    st.code(core.master_query(cfg), language="text")
    st.dataframe(pd.DataFrame(
        [{"ПН": r["pn_code"], "Направление": r["pn_name"],
          "areas": len(mapping.get(core.pn_code_to_matrix(r["pn_code"]), []))} for r in info_sel]),
        hide_index=True, use_container_width=True)

col_run, col_probe = st.columns(2)
run = col_run.button("▶ Run verification", type="primary", use_container_width=True)
probe = col_probe.button("🔎 Probe one record (schema check)", use_container_width=True)

def need_key():
    if not api_key:
        st.warning("Enter your WoS Expanded API key in the sidebar."); return False
    return True

if probe and need_key():
    with st.spinner("Fetching one record…"):
        try:
            recs, total = core.fetch_sample(cfg, api_key, core.master_query(cfg), n=1)
            if recs:
                st.write("**Extracted fields:**")
                st.json(core.probe_record(recs[0]))
                with st.expander("Raw record JSON"):
                    st.json(recs[0])
            else:
                st.warning("No records — check the organization string / years.")
        except Exception as e:
            st.error(str(e))

if run and need_key():
    bar = st.progress(0.0, text="Querying Web of Science…")
    def prog(got, total):
        bar.progress(min((got / total) if total else 1.0, 1.0),
                     text="Fetched %d / %s records" % (got, total))
    try:
        records, total = core.fetch_all(cfg, api_key, core.master_query(cfg), progress=prog)
    except Exception as e:
        st.error("Fetch failed: %s" % e); st.stop()
    parsed = core.parse_records(records)
    bar.progress(1.0, text="Parsed %d unique documents" % len(parsed))
    st.session_state.update(parsed=parsed,
                            results=core.verify_all(parsed, info_sel, keysets, cfg), cfg=cfg)

if "results" in st.session_state:
    parsed = st.session_state["parsed"]
    results = st.session_state["results"]
    rcfg = st.session_state["cfg"]
    st.subheader("Results")
    st.caption("Unique documents fetched: %d  ·  org: %s  ·  years: %d–%d" %
               (len(parsed), rcfg.org_query, rcfg.py_from, rcfg.py_to))

    rows, tot, match = [], 0, 0
    for r in results:
        row = {"ПН": r["pn"], "Направление": r["name"], "Docs": r["n_documents"]}
        if r["cmp"]:
            for key in core.METRIC_KEYS:
                v = r["cmp"][key]
                row["%s file" % key] = v["expected"]
                row["%s calc" % key] = v["computed"]
                if v["match"]:
                    row["%s Δ" % key] = "✅"
                elif v["direction"] == "higher":
                    row["%s Δ" % key] = "🔼 +%g" % v["delta"]
                elif v["direction"] == "lower":
                    row["%s Δ" % key] = "🔽 %g" % v["delta"]
                else:
                    row["%s Δ" % key] = "—"
                tot += 1; match += 1 if v["match"] else 0
        else:
            row["note"] = r.get("note", "")
        rows.append(row)
    if tot:
        st.metric("Indicators matching", "%d / %d" % (match, tot))
    df = pd.DataFrame(rows)
    def _hl(col):
        if col.name.endswith("Δ"):
            out = []
            for val in col:
                s = str(val)
                if s.startswith("✅"):
                    out.append("background-color:#C6EFCE")
                elif s.startswith("🔼") or s.startswith("🔽"):
                    out.append("background-color:#FFC7CE")
                else:
                    out.append("")
            return out
        return ["" for _ in col]
    st.dataframe(df.style.apply(_hl, axis=0), hide_index=True, use_container_width=True)
    st.download_button("⬇ Download report (.xlsx)",
                       data=core.report_to_xlsx_bytes(results, rcfg),
                       file_name="wos_verification_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("Data diagnostics"):
        dt = Counter(t for d in parsed for t in d["doctypes"])
        st.write("**Document types:**", dict(dt.most_common()))
        st.write("**Docs matching no mapped area (ignored):**",
                 sum(1 for d in parsed if not d["area_keys"]))
