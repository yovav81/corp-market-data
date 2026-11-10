#!/usr/bin/env python3
"""check_gov2.py — בדיקת קובץ הממשלתי v2: חישוב מחדש ב-LibreOffice, ספירת שגיאות נוסחה, זהות מול הפייתון.
שימוש: python gov/check_gov2.py <ממשלתי.xlsx> <data>   → יוצא 1 אם יש שגיאה או זהות שנכשלה.
"""
from __future__ import annotations
import os, subprocess, sys, tempfile
from pathlib import Path
import numpy as np, pandas as pd, openpyxl

HERE = Path(__file__).resolve().parent; sys.path.insert(0, str(HERE))
TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]


def recalc(xlsx: Path | str) -> Path:
    """ממיר ב-LibreOffice headless (חישוב מלא) ומחזיר את הנתיב לעותק המחושב."""
    xlsx = Path(xlsx); tmp = Path(tempfile.mkdtemp(prefix="gov2_recalc_"))
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp), str(xlsx)], check=True, capture_output=True, timeout=900)
    return tmp / xlsx.name


def count_errors(xlsx_recalced: Path | str) -> dict[str, int]:
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    for ws in wb.worksheets:
        n = 0
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and v.startswith("#") and v.rstrip("!?/0ADEFILMNRUV!") != v: n += 1
                elif isinstance(v, str) and v in ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "Err:502", "Err:504", "Err:508", "Err:511"): n += 1
        if n: out[ws.title] = n
    return out


def parity_curves(xlsx_recalced: Path | str, data: Path | str) -> float:
    """מקס' הפרש (נ"ב) בין העקום באקסל (03 B/C) לעקום הפייתון (curves_latest.csv)."""
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); ws = wb["03_עקומים"]
    cur = pd.read_csv(Path(data) / "gov" / "curves_latest.csv", encoding="utf-8-sig").set_index("tenor")
    worst = 0.0
    for i, t in enumerate(TENORS):
        r = 6 + i
        for col, key in (("B", "nominal"), ("C", "real")):
            v = ws[f"{col}{r}"].value
            if isinstance(v, (int, float)) and not np.isnan(cur.loc[t, key]):
                worst = max(worst, abs(float(v) - float(cur.loc[t, key])) * 100)
    return round(worst, 3)


def named_value(wb, name: str):
    """ערך תא לפי שם מוגדר (חוברת מחושבת, data_only)."""
    ref = wb.defined_names[name].attr_text
    sh, cell = ref.rsplit("!", 1); ws = wb[sh.strip("'")]
    return ws[cell.replace("$", "")].value


def parity_decisions(xlsx_recalced: Path | str, data: Path | str, s: dict | None = None) -> dict:
    """אקסל מול פייתון: משקל צמוד (±0.001), מח״מ שקלי/צמוד/משוקלל (±0.01). מחזיר {key: (excel, python, ok)}."""
    import gov_decide
    s = s or gov_decide.load_settings()
    py = gov_decide.decide(data, s)
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    def near(a, b, tol):
        try: return abs(float(a) - float(b)) <= tol
        except (TypeError, ValueError): return (a in (None, "")) and (b in (None, ""))
    for key, nm, tol in (("linked_weight", "m_linked", 0.001), ("dur_nominal", "m_dur_nominal", 0.01), ("dur_real", "m_dur_real", 0.01), ("dur_weighted", "m_dur_weighted", 0.01)):
        xl = named_value(wb, nm); out[key] = (xl, py.get(key), near(xl, py.get(key), tol))
    return out


def parity_optimizer(xlsx_recalced: Path | str, data: Path | str, s: dict | None = None) -> dict:
    """אקסל (13) מול gov_decide.optimize: הסתברויות (±1e-4 — העקום באקסל מחושב מהניירות, בפייתון מה-CSV), משקל (±0.001), מח״מים, תשואה צפויה ובתרחיש הגרוע (±0.005), סטטוס."""
    import gov_decide
    s = s or gov_decide.load_settings(); g = Path(data) / "gov"
    curve = pd.read_csv(g / "curves_latest.csv", encoding="utf-8-sig")
    sg = pd.read_csv(g / "signals_latest.csv", encoding="utf-8-sig") if (g / "signals_latest.csv").exists() else pd.DataFrame()
    sig = dict(zip(sg["key"], sg["val"])) if len(sg) else {}
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    h = pd.read_csv(g / "decisions_history.csv", encoding="utf-8-sig") if (g / "decisions_history.csv").exists() else pd.DataFrame()
    date = str(curve["date"].iloc[0]); prev = None
    if len(h) and all(c in h.columns for c in ("linked_weight", "opt_leg_n", "opt_leg_r")):
        y = h[h["date"] < date].sort_values("date").tail(1)
        if len(y): prev = tuple(float(y[c].iloc[0]) if pd.notna(y[c].iloc[0]) else float("nan") for c in ("linked_weight", "opt_leg_n", "opt_leg_r"))
    fc = named_value(wb, "infl_fc"); fc = None if fc in (None, "") else float(fc)
    o = gov_decide.optimize(curve, s, sig, float(named_value(wb, "r_linked")), fc, prev)
    ref = wb.defined_names["o_probs"].attr_text; sh, rng = ref.rsplit("!", 1)
    xp = [float(c[0].value) for c in wb[sh.strip("'")][rng.replace("$", "")]]
    out["probs"] = ([round(x, 4) for x in xp], [round(x, 4) for x in o["probs"]], max(abs(a - b) for a, b in zip(xp, o["probs"])) <= 1e-4)
    def near(a, b, tol):
        try: return abs(float(a) - float(b)) <= tol
        except (TypeError, ValueError): return (a in (None, "")) and (b in (None, ""))
    for nm, key, tol in (("o_linked", "w", 0.001), ("o_dur_nominal", "dur_n", 0.005), ("o_dur_real", "dur_r", 0.005), ("o_exp", "exp", 0.005), ("o_worst", "worst", 0.005)):
        xl = named_value(wb, nm); out[nm] = (xl, o.get(key), near(xl, o.get(key), tol))
    out["o_status"] = (named_value(wb, "o_status"), o["status"], named_value(wb, "o_status") == o["status"])
    out["o_legs"] = ((named_value(wb, "o_lab_nominal"), named_value(wb, "o_lab_real")), (o.get("lab_n"), o.get("lab_r")),
                     str(named_value(wb, "o_lab_nominal")) == str(o.get("lab_n")) and str(named_value(wb, "o_lab_real")) == str(o.get("lab_r")))
    return out


def parity_structures(xlsx_recalced: Path | str, data: Path | str, s: dict | None = None) -> dict:
    """אקסל מול פייתון: סוג המבנה, המפתחות והמשקלים (±0.005) לשני האפיקים, ביעד המח״מ שבאקסל."""
    import gov_decide
    s = s or gov_decide.load_settings(); g = Path(data) / "gov"
    bonds = pd.read_csv(g / "bonds_latest.csv", encoding="utf-8-sig"); curve = pd.read_csv(g / "curves_latest.csv", encoding="utf-8-sig")
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    for kind in ("nominal", "real"):
        tgt = named_value(wb, f"u_dur_{kind}")
        py = gov_decide.best_structure(bonds, kind, float(tgt), curve, s)
        xl_kind = named_value(wb, f"m_struct_{kind}_verdict")
        ids = [str(named_value(wb, f"s_{kind}_id{n}")) for n in (1, 2) if named_value(wb, f"s_{kind}_id{n}") not in (None, "")]
        ws_ = [float(named_value(wb, f"s_{kind}_w{n}")) for n in (1, 2) if named_value(wb, f"s_{kind}_w{n}") not in (None, "")]
        py_ids = [str(i) for i in py["ids"]]; py_w = [float(w) for w in py["weights"]]
        ok = xl_kind == py["kind"] and ids == py_ids and len(ws_) == len(py_w) and all(abs(a - b) <= 0.005 for a, b in zip(ws_, py_w))
        out[f"struct_{kind}"] = ((xl_kind, ids, ws_), (py["kind"], py_ids, py_w), ok)
    return out


def parity_top5(xlsx_recalced: Path | str, data: Path | str, s: dict | None = None) -> dict:
    import gov_decide
    s = s or gov_decide.load_settings(); g = Path(data) / "gov"
    bonds = pd.read_csv(g / "bonds_latest.csv", encoding="utf-8-sig"); curve = pd.read_csv(g / "curves_latest.csv", encoding="utf-8-sig")
    bh = pd.read_csv(g / "bond_history.csv", encoding="utf-8-sig") if (g / "bond_history.csv").exists() else pd.DataFrame(columns=["key", "kind", "spread_bp"])
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    for kind in ("nominal", "real"):
        ref = wb.defined_names[f"top5_{kind}"].attr_text; sh, rng = ref.rsplit("!", 1); ws = wb[sh.strip("'")]
        xl = [str(c[0].value) for c in ws[rng.replace("$", "")] if c[0].value not in (None, "")]
        py = [str(k) for k in gov_decide.series_rank(bonds, bh, kind, curve, s).head(5)["key"]]
        out[f"top5_{kind}"] = (xl, py, xl == py)
    return out


def scenario_table(wb) -> dict:
    """{(scenario_key, portfolio): value} מהשמות scen_<key>_<port> בחוברת מחושבת."""
    out = {}
    for nm in wb.defined_names:
        if nm.startswith("scen_"):
            body = nm[5:]; port = body.rsplit("_", 1)[1]; key = body[: -(len(port) + 1)]
            out[(key, port)] = named_value(wb, nm)
    return out


def parity_foreign(xlsx_recalced: Path | str, data: Path | str, s: dict | None = None) -> dict:
    import gov_decide
    s = s or gov_decide.load_settings(); g = Path(data) / "gov"
    curve = pd.read_csv(g / "curves_latest.csv", encoding="utf-8-sig")
    fo = pd.read_csv(g / "foreign_latest.csv", encoding="utf-8-sig") if (g / "foreign_latest.csv").exists() else pd.DataFrame()
    wb = openpyxl.load_workbook(xlsx_recalced, data_only=True); out = {}
    fv = gov_decide.foreign_verdict(fo, curve, s) if len(fo) else pd.DataFrame()
    for t, key in ((5, "foreign_5y"), (10, "foreign_10y")):
        xl = named_value(wb, f"f_verdict_{t}y")
        py = fv[fv["tenor"] == t]["verdict"].iloc[0] if len(fv) and (fv["tenor"] == t).any() else "אין נתון"
        out[key] = (xl, py, xl == py)
    return out


def main(xlsx: str, data: str) -> int:
    rc = recalc(xlsx); errs = count_errors(rc); worst = parity_curves(rc, data)
    print("שגיאות נוסחה:", errs or "0"); print("זהות עקומים (מקס' נ״ב):", worst)
    ok = not errs and worst <= 0.5
    for fn in ("parity_decisions", "parity_optimizer", "parity_structures", "parity_top5", "parity_foreign"):
        f = globals().get(fn)
        if f:
            res = f(rc, data); print(fn + ":", res); ok = ok and all(v[2] for v in res.values())
    print("תקין" if ok else "נכשל"); return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
