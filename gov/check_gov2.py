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


def main(xlsx: str, data: str) -> int:
    rc = recalc(xlsx); errs = count_errors(rc); worst = parity_curves(rc, data)
    print("שגיאות נוסחה:", errs or "0"); print("זהות עקומים (מקס' נ״ב):", worst)
    ok = not errs and worst <= 0.5
    for fn in ("parity_decisions", "parity_structures", "parity_top5", "parity_foreign"):
        f = globals().get(fn)
        if f:
            res = f(rc, data); print(fn + ":", res); ok = ok and all(v[2] for v in res.values())
    print("תקין" if ok else "נכשל"); return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
