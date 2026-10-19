"""Daily Maya watcher for the CORP issuers.

    python maya/watch.py --out maya_out            # new reports since the last check
    python maya/watch.py --out maya_out --init     # resolve ids + baseline (no downloads)
    python maya/watch.py --out maya_out --since 2026-10-01 --ignore-seen   # test / re-pull a window

State lives in maya/state.json ({companyId: {"last_check": date, "seen": {reportId: publishDate}}})
and is committed back by the workflow. Every report found in the window and not yet seen is
recorded as seen (kept or skipped); kept ones are downloaded to <out>/<cid>_<name>/ with a
catalog.csv, and <out>/summary.json lists everything new per issuer.
"""
import argparse, csv, datetime as dt, json, pathlib, sys, time
from urllib.parse import quote
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from fetch import Maya, decide, clean, MAX_FILE, B   # noqa: E402

HERE = pathlib.Path(__file__).parent
ISSUERS = HERE / "issuers.json"
STATE = HERE / "state.json"
LOOKBACK_DAYS = 5          # overlap with the previous check, so late-indexed reports are not missed
KEEP_SEEN_DAYS = 180

def load(p, default):
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default

def save(p, obj):
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")

def list_reports(m, cid, frm, to):
    out, off = [], 0
    while True:
        page = m.api("/api/v1/reports/companies", {"pageNumber": 1, "limit": 30, "companyId": cid, "offset": off,
                                                   "fromDate": frm.isoformat(), "toDate": to.isoformat()})
        out.extend(page)
        if len(page) < 30:
            return out
        off += 30

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="maya_out")
    ap.add_argument("--init", action="store_true")
    ap.add_argument("--since", default="")
    ap.add_argument("--ignore-seen", action="store_true")
    ap.add_argument("--only", default="", help="comma-separated issuer keys")
    a = ap.parse_args()
    today = dt.date.today()
    issuers = load(ISSUERS, [])
    state = load(STATE, {})
    out = pathlib.Path(a.out); out.mkdir(parents=True, exist_ok=True)
    summary = {"run_date": today.isoformat(), "mode": "init" if a.init else "watch", "issuers": []}
    with sync_playwright() as p:
        m = Maya(p)
        for iss in issuers:
            if a.only and iss["key"] not in a.only.split(","):
                continue
            if not iss.get("maya_id"):
                res = m.api(f"/api/v1/companies/autocomplete?search={quote(iss['maya_name'])}&take=10")
                comps = [r for r in res if r.get("type") == "COMPANY"]
                exact = [r for r in comps if iss["maya_name"] in (r.get("value") or "")]
                if not (exact or comps):
                    summary["issuers"].append({"key": iss["key"], "error": f"no Maya match for {iss['maya_name']}",
                                               "candidates": res}); continue
                pick = (exact or comps)[0]
                iss["maya_id"], iss["maya_long_name"] = int(pick["key"]), pick.get("value")
            cid = str(iss["maya_id"])
            st = state.setdefault(cid, {"last_check": None, "seen": {}})
            if a.since:
                frm = dt.date.fromisoformat(a.since)
            elif st["last_check"]:
                frm = dt.date.fromisoformat(st["last_check"]) - dt.timedelta(days=LOOKBACK_DAYS)
            else:
                frm = today - dt.timedelta(days=60)
            reps = list_reports(m, iss["maya_id"], frm, today)
            new = [r for r in reps if a.ignore_seen or str(r["id"]) not in st["seen"]]
            new.sort(key=lambda r: r["publishDate"])
            entry = {"key": iss["key"], "maya_id": iss["maya_id"], "corp_folder": iss["corp_folder"],
                     "name": iss["name"], "window": [frm.isoformat(), today.isoformat()],
                     "listed": len(reps), "new": len(new), "kept": []}
            root = out / f"{cid}_{clean(iss['name'], 30)}"
            rows = []
            for r in new:
                st["seen"][str(r["id"])] = r["publishDate"][:10]
                if a.init:
                    continue
                keep, why = decide(r)
                atts = [x for x in (r.get("attachments") or []) if x.get("url") and x.get("fileType") != "edit"]
                saved = []
                if keep:
                    root.mkdir(parents=True, exist_ok=True)
                    date = r["publishDate"][:10]
                    for k, x in enumerate(atts):
                        ext = x["url"].rsplit(".", 1)[-1].lower()
                        fn = (f"{date}_{r['id']}_{(r.get('formId') or 'NA')}_{clean(r['title'], 60)}"
                              + (f"__{k}" if len(atts) > 1 else "") + f".{ext}")
                        data = m.file(x["url"])
                        if data is None:
                            saved.append("FAILED"); continue
                        if len(data) > MAX_FILE:
                            saved.append(f"TOO_BIG:{len(data)}"); continue
                        (root / fn).write_bytes(data); saved.append(fn); time.sleep(0.4)
                    entry["kept"].append({"report_id": r["id"], "date": r["publishDate"][:16].replace("T", " "),
                                          "form": r.get("formId") or "", "title": r.get("title") or "",
                                          "reason": why, "files": saved, "maya_url": f"{B}/he/reports/{r['id']}"})
                rows.append({"report_id": r["id"], "date": r["publishDate"][:19], "form": r.get("formId") or "",
                             "title": r.get("title") or "", "keep": "כן" if keep else "לא", "reason": why,
                             "attachments": " | ".join(f"{x.get('fileType')}:{x.get('fileName') or ''}:{x['url']}" for x in atts),
                             "saved": " | ".join(saved), "maya_url": f"{B}/he/reports/{r['id']}"})
            if rows and entry["kept"]:
                with open(root / "catalog.csv", "w", newline="", encoding="utf-8-sig") as fh:
                    w = csv.DictWriter(fh, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
                entry["dir"] = root.name
            st["last_check"] = today.isoformat()
            cutoff = (today - dt.timedelta(days=KEEP_SEEN_DAYS)).isoformat()
            st["seen"] = {k: v for k, v in st["seen"].items() if v >= cutoff}
            summary["issuers"].append(entry)
            print(f"{iss['key']}: listed {len(reps)}, new {len(new)}, kept {len(entry['kept'])}", flush=True)
        m.b.close()
    save(ISSUERS, issuers)
    save(STATE, state)
    summary["total_kept"] = sum(len(e.get("kept", [])) for e in summary["issuers"])
    save(out / "summary.json", summary)
    print(json.dumps({k: v for k, v in summary.items() if k != "issuers"}, ensure_ascii=False))

if __name__ == "__main__":
    main()
