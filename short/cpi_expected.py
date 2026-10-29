"""
cpi_expected.py — המדד הידוע ותוואי המדדים הצפויים לחודשים הבאים (לחישוב התזרים הנומינלי של אג"ח צמודות קצרות).

מקורות (לא מהדוחות — מסומנים עם מקור):
  * מדד ידוע: הלמ״ס, API המדדים (סדרה 120010), לפי מועד הפרסום (15 לחודש העוקב).
  * ציפיות לשנה: בנק ישראל, מאגר הסדרות (SDMX, dataflow ZCM, חודשי):
        EXP_HAZ_MAD_AVG_01Y_MA   ממוצע החזאים ל-12 החודשים הבאים   (ברירת המחדל לתוואי)
        EXP_EXP_01Y_BI_MA        ציפיות משוק ההון לשנה (אג"ח ממשלתי)
        EXP_EXP_IRB_1Y_MA        ציפיות מריביות הבנקים לשנה
        EXP_EXP_ISSWIT1_01Y_MA   ציפיות מחוזי אינפלציה לשנה
  * עונתיות: ממוצע השינוי החודשי לכל חודש קלנדרי ב-SEASON_YEARS השנים המלאות האחרונות, מנוכה ממוצע (הלמ״ס).
  תוואי חודשי = עונתיות(חודש) + (1+ציפייה שנתית)^(1/12) − 1.
  * דריסה ידנית: data/short/cpi_manual.csv (עמודות: ym, pct, source, as_of) — אם קיימת שורה לחודש, היא גוברת.
  כשקובץ הממשלתי ייבנה, הציפיות מהעקום ייכנסו כאן דרך SHORT_EXP_SOURCE=gov או דרך cpi_manual.csv.

פלט: data/short/cpi_expected.csv — שורה לחודש (ידוע + 18 חודשים צפויים): ym, pub_date, kind, index, monthly_pct, cum_factor, source, as_of
       data/short/cpi_sources.json — הציפיות שנמשכו מבנק ישראל (לתיעוד בגיליון)
הרצה: python short/cpi_expected.py data
"""
from __future__ import annotations
import csv, datetime as dt, io, json, os, sys, urllib.request
from pathlib import Path

UA = {"User-Agent": "Mozilla/5.0 corp-market-data"}
CBS_ID = 120010
SEASON_YEARS = int(os.environ.get("SHORT_SEASON_YEARS", "6"))
HORIZON = int(os.environ.get("SHORT_HORIZON_MONTHS", "18"))
EXP_SOURCE = os.environ.get("SHORT_EXP_SOURCE", "forecasters")   # forecasters | market | banks | swaps | gov
BOI_SERIES = {"forecasters": ("EXP_HAZ_MAD_AVG_01Y_MA", "בנק ישראל — ממוצע החזאים ל-12 חודשים"),
              "market": ("EXP_EXP_01Y_BI_MA", "בנק ישראל — ציפיות משוק ההון לשנה"),
              "banks": ("EXP_EXP_IRB_1Y_MA", "בנק ישראל — ציפיות מריביות הבנקים לשנה"),
              "swaps": ("EXP_EXP_ISSWIT1_01Y_MA", "בנק ישראל — ציפיות מחוזי אינפלציה לשנה")}


def get(url, timeout=60):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8-sig", errors="replace")


def pub_date(y, m):
    """מועד פרסום המדד לחודש (y,m): ה-15 לחודש העוקב."""
    m2 = m + 1; y2 = y
    if m2 > 12: m2 = 1; y2 += 1
    return dt.date(y2, m2, 15)


def fetch_cbs(start_year=2010):
    """{(שנה,חודש): {"value": מדד (בסיס נוכחי), "base": תיאור הבסיס, "pct": שינוי חודשי % מה-API אם קיים}}"""
    txt = get(f"https://api.cbs.gov.il/index/data/price?id={CBS_ID}&format=json&download=false&startPeriod=01-{start_year}&endPeriod=12-{dt.date.today().year}&lang=he")
    js = json.loads(txt); out = {}
    def walk(o):
        if isinstance(o, dict):
            if "year" in o and "month" in o and isinstance(o.get("currBase"), dict) and o["currBase"].get("value") is not None:
                pct = None
                for k in ("percent", "percentChange", "change"):
                    v = o.get(k)
                    if isinstance(v, dict):
                        for kk in ("monthly", "month", "m", "value"):
                            if v.get(kk) is not None:
                                try: pct = float(v[kk]); break
                                except (TypeError, ValueError): pass
                    elif v is not None:
                        try: pct = float(v)
                        except (TypeError, ValueError): pass
                    if pct is not None: break
                out[(int(o["year"]), int(o["month"]))] = {"value": float(o["currBase"]["value"]),
                                                          "base": str(o["currBase"].get("baseDesc") or o["currBase"].get("base") or ""), "pct": pct}
            for v in o.values(): walk(v)
        elif isinstance(o, list):
            for v in o: walk(v)
    walk(js)
    return out


def monthly_change(cpi, k):
    """השינוי החודשי (שבר) לחודש k: מה-API אם קיים, אחרת יחס למדד הקודם באותו בסיס; None במעבר בסיס."""
    o = cpi.get(k)
    if not o: return None
    if o.get("pct") is not None and abs(o["pct"]) < 10: return o["pct"] / 100
    pk = (k[0], k[1] - 1) if k[1] > 1 else (k[0] - 1, 12)
    p = cpi.get(pk)
    if not p or p["base"] != o["base"]: return None
    return o["value"] / p["value"] - 1


def fetch_boi() -> dict:
    """{key: {"series":..., "name":..., "period": "YYYY-MM", "value": pct}}"""
    start = (dt.date.today() - dt.timedelta(days=200)).isoformat()
    txt = get(f"https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/ZCM/1.0?startperiod={start}&format=csv")
    rows = list(csv.DictReader(io.StringIO(txt)))
    out = {}
    for key, (code, name) in BOI_SERIES.items():
        obs = [r for r in rows if r.get("SERIES_CODE") == code and r.get("OBS_VALUE") not in (None, "")]
        if not obs: continue
        last = max(obs, key=lambda r: r["TIME_PERIOD"])
        out[key] = {"series": code, "name": name, "period": last["TIME_PERIOD"], "value": float(last["OBS_VALUE"])}
    return out


def seasonality(cpi: dict, years: int):
    """סטיית השינוי החודשי מהממוצע לכל חודש קלנדרי (שבר), על השנים המלאות האחרונות.
    שינויים שלא ניתן לחשב (מעבר בסיס) או חריגים (|שינוי| > 1.5%) לא נכנסים לממוצע."""
    ys = sorted({y for y, _ in cpi})
    full = [y for y in ys if all((y, m) in cpi for m in range(1, 13))][-years:]
    ch = {m: [] for m in range(1, 13)}
    for y in full:
        for m in range(1, 13):
            c = monthly_change(cpi, (y, m))
            if c is not None and abs(c) <= 0.015: ch[m].append(c)
    avg = {m: sum(v) / len(v) for m, v in ch.items() if v}
    mean = sum(avg.values()) / len(avg) if avg else 0.0
    return {m: v - mean for m, v in avg.items()}, full


def load_manual(p: Path) -> dict[tuple[int, int], dict]:
    out = {}
    if p.exists():
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            try:
                y, m = r["ym"].split("-"); out[(int(y), int(m))] = {"pct": float(r["pct"]) / 100, "source": r.get("source", "ידני"), "as_of": r.get("as_of", "")}
            except Exception:
                pass
    return out


def build(cpi, boi, manual, today: dt.date):
    known = max(k for k in cpi if pub_date(*k) <= today - dt.timedelta(days=0))
    # ידוע = פורסם; המדד שפורסם היום ב-15 נחשב ידוע ממחר — לכן התנאי הוא pub_date < today
    known = max(k for k in cpi if pub_date(*k) < today)
    seas, seas_years = seasonality(cpi, SEASON_YEARS)
    exp = boi.get(EXP_SOURCE) or boi.get("forecasters")
    annual = exp["value"] / 100 if exp else None
    mo = (1 + annual) ** (1 / 12) - 1 if annual is not None else None
    rows = []
    # היסטוריה קצרה (12 חודשים ידועים) לצורך הצגה
    hist = sorted(k for k in cpi if k <= known)[-13:]
    for i, k in enumerate(hist):
        c = monthly_change(cpi, k)
        rows.append({"ym": f"{k[0]}-{k[1]:02d}", "pub_date": pub_date(*k).isoformat(), "kind": "ידוע", "index": cpi[k]["value"],
                     "monthly_pct": round(c * 100, 3) if c is not None else None, "cum_factor": None,
                     "source": "הלמ״ס — API המדדים", "as_of": today.isoformat()})
    f = 1.0; y, m = known
    for _ in range(HORIZON):
        m += 1
        if m > 12: m = 1; y += 1
        if (y, m) in manual:
            e, src = manual[(y, m)]["pct"], manual[(y, m)]["source"]; kind = "צפוי (ידני)"
        elif mo is not None:
            e = seas.get(m, 0.0) + mo; kind = "צפוי"
            src = f"{exp['name']} ({exp['period']}: {exp['value']:.2f}%) + עונתיות הלמ״ס {seas_years[0]}–{seas_years[-1]}"
        else:
            e = 0.0; kind = "צפוי (אין מקור)"; src = "אין ציפיות — 0"
        f *= 1 + e
        rows.append({"ym": f"{y}-{m:02d}", "pub_date": pub_date(y, m).isoformat(), "kind": kind, "index": None,
                     "monthly_pct": round(e * 100, 3), "cum_factor": round(f, 6), "source": src, "as_of": today.isoformat()})
    # מקדם מצטבר יחסית למדד הידוע (ידוע = 1)
    for r in rows:
        if r["kind"] == "ידוע": r["cum_factor"] = 1.0 if r["ym"] == f"{known[0]}-{known[1]:02d}" else None
    return rows, known


def main(data="data"):
    d = Path(data) / "short"; d.mkdir(parents=True, exist_ok=True)
    today = dt.date.today()
    cpi = fetch_cbs()
    try:
        boi = fetch_boi()
    except Exception as e:
        print("אזהרה: בנק ישראל לא זמין —", e); boi = {}
    manual = load_manual(d / "cpi_manual.csv")
    rows, known = build(cpi, boi, manual, today)
    with open(d / "cpi_expected.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["ym", "pub_date", "kind", "index", "monthly_pct", "cum_factor", "source", "as_of"])
        w.writeheader(); w.writerows(rows)
    json.dump({"as_of": today.isoformat(), "known_month": f"{known[0]}-{known[1]:02d}", "known_index": cpi[known]["value"],
               "known_base": cpi[known]["base"], "seasonality_pct": {m: round(v * 100, 3) for m, v in seasonality(cpi, SEASON_YEARS)[0].items()},
               "expectations": boi, "used": EXP_SOURCE, "season_years": SEASON_YEARS},
              open(d / "cpi_sources.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"מדד ידוע: {known[0]}-{known[1]:02d} = {cpi[known]['value']} ({cpi[known]['base']}); ציפיות: " + ", ".join(f"{k}={v['value']:.2f}% ({v['period']})" for k, v in boi.items()))
    for r in rows:
        if r["kind"] != "ידוע": print(r["ym"], r["monthly_pct"], r["cum_factor"], r["kind"])


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data")
