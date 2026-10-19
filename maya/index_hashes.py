"""reports_index.csv (written by the PC sync for each issuer) -> "md5  name" lines for sort_to_corp.py --existing.

    python maya/index_hashes.py <corp-data>/02_Issuers/<folder>/_log/reports_index.csv > hashes.txt
"""
import csv, sys

if len(sys.argv) > 1:
    try:
        for r in csv.DictReader(open(sys.argv[1], encoding="utf-8-sig")):
            if r.get("md5"):
                print(f"{r['md5']}  {r['path']}")
    except FileNotFoundError:
        pass
