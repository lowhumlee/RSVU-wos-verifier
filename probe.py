#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe.py — confirm the Expanded API field paths on YOUR subscription.

Run locally (key from environment, never hard-coded):

    export WOS_API_KEY="xxxx"
    python probe.py                          # 2 records from the master query
    python probe.py --n 3
    python probe.py --ut WOS:000XXXXXXXXXXXX  # a specific record
    python probe.py --org 'OG=(Medical University Varna)' --from 2021 --to 2025
    python probe.py --raw                     # also dump the full raw JSON

It prints, per record, what core.py's extractors pull (times-cited from the WOS
silo, doctypes, ascatype='extended' research areas, #institutions). Paste the
output back so we can confirm or adjust the extractors before the full run.
"""
import argparse
import json
import os
import sys

import core


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=2)
    ap.add_argument("--ut", default=None, help="fetch one specific record by UT")
    ap.add_argument("--org", default="OG=(Medical University Varna)")
    ap.add_argument("--from", dest="yfrom", type=int, default=2021)
    ap.add_argument("--to", dest="yto", type=int, default=2025)
    ap.add_argument("--raw", action="store_true", help="also print the full raw JSON")
    args = ap.parse_args()

    api_key = os.environ.get("WOS_API_KEY")
    if not api_key:
        sys.exit("Set WOS_API_KEY in your environment first (do not paste it anywhere).")

    cfg = core.Config(org_query=args.org, py_from=args.yfrom, py_to=args.yto)
    query = "UT=(%s)" % args.ut if args.ut else core.master_query(cfg)
    print("Query:", query)

    recs, total = core.fetch_sample(cfg, api_key, query, n=args.n)
    print("RecordsFound:", total, "| fetched:", len(recs))
    print("=" * 70)
    for i, rec in enumerate(recs, 1):
        print("\n--- RECORD %d ---" % i)
        print(json.dumps(core.probe_record(rec), ensure_ascii=False, indent=2))
        if args.raw:
            print("\nRAW:")
            print(json.dumps(rec, ensure_ascii=False, indent=2)[:9000])
    print("\nIf any extracted field is empty/None/wrong, paste a RAW record "
          "(python probe.py --raw) so the extractor path can be corrected.")


if __name__ == "__main__":
    main()
