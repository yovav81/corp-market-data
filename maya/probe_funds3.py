"""Probe 3: header/structure of the monthly reports (ק203) that failed to normalize (Altshuler, Kesem)."""
import json, sys, pathlib, requests, collections
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
FILES = {"altshuler_08": "rtxt/1771001-1772000/T1771386-00.txt", "kesem_08": "rtxt/1770001-1771000/T1770573-00.txt", "kesem_07": "rtxt/1763001-1764000/T1763488-00.txt"}
log = {}
for k, u in FILES.items():
    r = requests.get("https://mayafiles.tase.co.il/" + u, headers={"User-Agent": UA}, timeout=180)
    raw = r.content
    for enc in ("utf-8-sig", "utf-8", "cp1255"):
        try:
            txt = raw.decode(enc); break
        except UnicodeDecodeError:
            txt = None
    lines = txt.splitlines()
    hdr = lines[0].split("\t")
    rows = [l.split("\t") for l in lines[1:]]
    ncols = collections.Counter(len(r) for r in rows)
    # duplicates of fund-level rows (asset type 301) per fund
    try:
        it, ia = hdr.index("סוג נכס"), hdr.index("מספר קרן")
        c301 = collections.Counter(r[ia].strip() for r in rows if len(r) > it and r[it].strip() == "301")
        dups = {f: n for f, n in c301.items() if n > 1}
    except ValueError:
        dups = "no cols"
    log[k] = {"status": r.status_code, "bytes": len(raw), "encoding": enc, "nlines": len(lines), "header": hdr, "ncols": dict(ncols),
              "first": lines[1][:600] if len(lines) > 1 else "", "dup301": dups if isinstance(dups, str) else dict(list(dups.items())[:10]),
              "ndup301": None if isinstance(dups, str) else len(dups)}
    print(k, json.dumps({a: b for a, b in log[k].items() if a != "first"}, ensure_ascii=False)[:1500], flush=True)
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
