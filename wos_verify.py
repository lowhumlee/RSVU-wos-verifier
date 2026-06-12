#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wos_verify.py — CLI runner via the WoS Expanded API.

  export WOS_API_KEY="xxxx"
  python wos_verify.py --selftest
  python wos_verify.py --probe                 # one raw record + extracted fields
  python wos_verify.py                          # all fields -> report.xlsx
  python wos_verify.py --pn 704 --collab-min 2
"""
import argparse
import json
import os
import sys

import core

DATA = os.path.join(os.path.dirname(__file__), "data")


def main():
    ap = argparse.ArgumentParser(description="Verify РСВУ WoS indicators via WoS Expanded API")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--probe", action="store_true", help="dump one record + extracted fields")
    ap.add_argument("--raw", action="store_true", help="with --probe, also print raw JSON")
    ap.add_argument("--pn", type=int, default=None)
    ap.add_argument("--org", default="OG=(Medical University Varna)")
    ap.add_argument("--from", dest="yfrom", type=int, default=2021)
    ap.add_argument("--to", dest="yto", type=int, default=2025)
    ap.add_argument("--collab-min", type=int, default=2)
    ap.add_argument("--count-reviews", action="store_true")
    ap.add_argument("--info", default=os.path.join(DATA, "Информация_ПН_ВУ_391003.xlsx"))
    ap.add_argument("--matrix", default=os.path.join(DATA, "WoS_Area_pn_MUV.xlsx"))
    ap.add_argument("--out", default="wos_verification_report.xlsx")
    args = ap.parse_args()

    if args.selftest:
        print("self-test:", core._selftest()); return

    cfg = core.Config(org_query=args.org, py_from=args.yfrom, py_to=args.yto,
                      collab_min_institutions=args.collab_min,
                      count_review_as_article=args.count_reviews)
    api_key = os.environ.get("WOS_API_KEY")
    if not api_key:
        sys.exit("Set WOS_API_KEY in your environment first.")

    if args.probe:
        recs, total = core.fetch_sample(cfg, api_key, core.master_query(cfg), n=2)
        print("RecordsFound:", total)
        for r in recs:
            print(json.dumps(core.probe_record(r), ensure_ascii=False, indent=2))
            if args.raw:
                print(json.dumps(r, ensure_ascii=False, indent=2)[:9000])
        return

    info = core.load_info(args.info)
    mapping, keysets = core.load_area_matrix(args.matrix)
    if args.pn:
        info = [r for r in info if r["pn_code"] == args.pn]
        if not info:
            sys.exit("ПН %s not in values file" % args.pn)

    print("MASTER query:", core.master_query(cfg))
    def prog(got, total):
        sys.stdout.write("\r  fetched %d/%s" % (got, total)); sys.stdout.flush()
    records, total = core.fetch_all(cfg, api_key, core.master_query(cfg), progress=prog)
    parsed = core.parse_records(records)
    print("\nunique documents:", len(parsed))

    results = core.verify_all(parsed, info, keysets, cfg)
    tot = match = 0
    for r in results:
        if not r["cmp"]:
            print("ПН%s %s -> %s" % (r["pn"], r["name"], r.get("note"))); continue
        print("\n=== ПН%s %s  (%d docs) ===" % (r["pn"], r["name"], r["n_documents"]))
        for k, v in r["cmp"].items():
            tot += 1; match += 1 if v["match"] else 0
            print("  %s %-14s file=%-8s calc=%-8s" %
                  ("OK " if v["match"] else "XX ", k, v["expected"], v["computed"]))
    print("\nMATCHED %d / %d" % (match, tot))
    with open(args.out, "wb") as f:
        f.write(core.report_to_xlsx_bytes(results, cfg))
    print("Report:", args.out)


if __name__ == "__main__":
    main()
