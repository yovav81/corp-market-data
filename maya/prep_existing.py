"""List issuer-folder files that could be duplicates of freshly fetched files (same size).

    python maya/prep_existing.py <device_listing.json> <fetch_dir> <device_issuer_root>

<device_listing.json>: the JSON returned by device_list_dir(<device_issuer_root>\\דוחות, recursive=true)
(the {"entries": [...]} object). Prints a JSON list of absolute device paths to stage; stage them,
run maya/hash_existing.py on them and pass "md5  name" lines to sort_to_corp.py --existing.
"""
import json, pathlib, sys

listing = json.load(open(sys.argv[1], encoding="utf-8"))
fetch = pathlib.Path(sys.argv[2])
root = sys.argv[3].rstrip("\\")
files = [p for p in fetch.iterdir() if p.is_file() and p.name != "catalog.csv"]
sizes = {p.stat().st_size for p in files}
htm_sizes = [p.stat().st_size for p in files if p.suffix.lower() in (".htm", ".html")]

def match(e):
    s, n = e.get("size"), e["name"].lower()
    if s in sizes:
        return True
    return n.endswith((".htm", ".html")) and any(abs(s - h) <= 0.03 * h for h in htm_sizes)

cands = [root + "\\דוחות\\" + e["name"].replace("/", "\\") for e in listing.get("entries", [])
         if e.get("type") == "file" and match(e)]
print(json.dumps(cands, ensure_ascii=False))
