"""Probe: fund managers' monthly holdings reports (טופס ק203 "דוח חודשי") on Maya.
Downloads one known report TXT (Meitav, August 2026) and records its structure, plus the fund-manager
report list that the Maya fund page uses, so the pipeline can be built on it."""
import json, os, sys, time, pathlib, io
import requests

OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out")
OUT.mkdir(parents=True, exist_ok=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
H = {"User-Agent": UA, "Accept-Language": "he-IL"}
log = []

def note(**k):
    k["t"] = time.strftime("%H:%M:%S"); log.append(k); print(json.dumps(k, ensure_ascii=False)[:500], flush=True)

# 1. the report TXT itself (public file on mayafiles)
URL = "https://mayafiles.tase.co.il/rtxt/1772001-1773000/T1772830-00.txt"
try:
    r = requests.get(URL, headers=H, timeout=120)
    note(kind="txt", status=r.status_code, ct=r.headers.get("content-type"), len=len(r.content))
    if r.ok:
        raw = r.content
        (OUT / "T1772830-00.txt").write_bytes(raw)
        for enc in ("utf-8-sig", "utf-8", "cp1255", "iso-8859-8"):
            try:
                txt = raw.decode(enc); break
            except UnicodeDecodeError:
                txt = None
        if txt:
            lines = txt.splitlines()
            note(kind="txt_struct", encoding=enc, nlines=len(lines), first=lines[:8], mid=lines[len(lines)//2:len(lines)//2+3], last=lines[-3:])
            seps = {s: lines[1].count(s) if len(lines) > 1 else 0 for s in ("|", "\t", ",", ";")}
            note(kind="seps", seps=seps)
except Exception as e:
    note(kind="txt_err", err=str(e)[:300])

# 2. Maya API: report metadata + the list of reports of this fund manager (form ק203 = דוח חודשי)
A = {"User-Agent": UA, "Accept": "application/json", "Accept-Language": "he-IL", "Referer": "https://maya.tase.co.il/he/funds/mutual-funds/5123138"}
for u in ("https://maya.tase.co.il/api/v1/reports/1772830",
          "https://maya.tase.co.il/api/v1/funds/mutual/5123138",
          "https://maya.tase.co.il/api/v1/reports/1772830/siblings?offset=0&limit=30"):
    try:
        r = requests.get(u, headers=A, timeout=60)
        body = r.text
        note(kind="api", url=u, status=r.status_code, ct=r.headers.get("content-type"), len=len(body), head=body[:600])
        (OUT / (u.rsplit("/", 2)[-2] + "_" + u.rsplit("/", 1)[-1].split("?")[0] + ".json")).write_text(body, encoding="utf-8")
    except Exception as e:
        note(kind="api_err", url=u, err=str(e)[:300])

(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
print("done", len(log))
