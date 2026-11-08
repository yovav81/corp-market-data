# -*- coding: utf-8 -*-
"""composition.py — הרכב חודשי לכל קרן (אחוז משווי נקי לפי אפיק) מתוך קובץ האחזקות המנורמל של החודש.

    python funds/composition.py data            # בונה composition.csv לכל חודש שאין לו, ו-composition_all.csv (כל החודשים)

עמודות (שברים של השווי הנקי; חשיפות נגזרים בערך נקוב):
  מניות_ישראל (= ת״א35 + ת״א90 + יתר), מניות_חו״ל (= ארה״ב + אירופה + אחר), קונצרני_צמוד, קונצרני_שקלי, קונצרני_מט״ח, קונצרני_חו״ל,
  ממשלתי_צמוד, ממשלתי_שקלי, ממשלתי_חו״ל, מק״מ, מזומן_ופיקדונות, נגזרים, אחר, מח״מ_מחושב (אג״ח, משוקלל), + המדווח מ-funds.csv:
  חשיפה_מניות_מדווחת, חשיפה_מטח_מדווחת, מחמ_מדווח, שווי_נכסים, סיווג, מנהל.
"""
import glob, gzip, os, pathlib, sys
import pandas as pd

COLS = ["מניות_ישראל", "ת״א35", "ת״א90", "יתר", "מניות_חו״ל", "ארה״ב", "אירופה", "מניות_חו״ל_אחר", "קונצרני", "קונצרני_צמוד", "קונצרני_שקלי", "קונצרני_מט״ח", "קונצרני_חו״ל",
        "ממשלתי", "ממשלתי_צמוד", "ממשלתי_שקלי", "ממשלתי_חו״ל", "מק״מ", "מזומן_ופיקדונות", "נגזרים", "אחר", "מח״מ_מחושב"]


def composition(h, fdf):
    h = h.copy()
    for c in ("משקל", "מחמ_x_משקל", "משקל_עם_מחמ"):
        h[c] = pd.to_numeric(h[c], errors="coerce").fillna(0.0)
    il = h["אזור"].eq("ישראל"); eq = h["אפיק"].eq("מניות"); co = h["אפיק"].eq("קונצרני"); go = h["אפיק"].eq("ממשלתי")
    mk = h["סוג_מכשיר"].eq("מק״מ")
    def s(mask): return h.loc[mask].groupby("מס_קרן")["משקל"].sum()
    out = pd.DataFrame(index=sorted(h["מס_קרן"].unique()))
    out["ת״א35"] = s(eq & il & h["תת_סיווג"].eq("ת״א 35")); out["ת״א90"] = s(eq & il & h["תת_סיווג"].eq("ת״א 90")); out["יתר"] = s(eq & il & ~h["תת_סיווג"].isin(["ת״א 35", "ת״א 90"]))
    out["ארה״ב"] = s(eq & ~il & h["תת_סיווג"].eq("ארה״ב")); out["אירופה"] = s(eq & ~il & h["תת_סיווג"].eq("אירופה")); out["מניות_חו״ל_אחר"] = s(eq & ~il & ~h["תת_סיווג"].isin(["ארה״ב", "אירופה"]))
    out["קונצרני_צמוד"] = s(co & il & h["הצמדה"].eq("צמוד מדד")); out["קונצרני_שקלי"] = s(co & il & h["הצמדה"].isin(["שקלי", "לא ידוע"])); out["קונצרני_מט״ח"] = s(co & il & h["הצמדה"].eq("מט״ח"))
    out["קונצרני_חו״ל"] = s(co & ~il)
    out["ממשלתי_צמוד"] = s(go & il & ~mk & h["הצמדה"].eq("צמוד מדד")); out["ממשלתי_שקלי"] = s(go & il & ~mk & ~h["הצמדה"].eq("צמוד מדד")); out["ממשלתי_חו״ל"] = s(go & ~il)
    out["מק״מ"] = s(mk); out["מזומן_ופיקדונות"] = s(h["אפיק"].eq("מזומן ופיקדונות")); out["נגזרים"] = s(h["אפיק"].eq("נגזרים")); out["אחר"] = s(h["אפיק"].eq("אחר"))
    out = out.fillna(0.0)
    out["מניות_ישראל"] = out["ת״א35"] + out["ת״א90"] + out["יתר"]; out["מניות_חו״ל"] = out["ארה״ב"] + out["אירופה"] + out["מניות_חו״ל_אחר"]
    out["קונצרני"] = out["קונצרני_צמוד"] + out["קונצרני_שקלי"] + out["קונצרני_מט״ח"] + out["קונצרני_חו״ל"]
    out["ממשלתי"] = out["ממשלתי_צמוד"] + out["ממשלתי_שקלי"] + out["ממשלתי_חו״ל"] + out["מק״מ"]
    bonds = h[(co | go) & (h["משקל_עם_מחמ"] > 0)]
    g = bonds.groupby("מס_קרן").agg(dw=("מחמ_x_משקל", "sum"), w=("משקל_עם_מחמ", "sum"))
    out["מח״מ_מחושב"] = (g["dw"] / g["w"]).reindex(out.index)
    out = out[COLS].round(5)
    f = fdf.copy(); f["מס_קרן"] = f["מס_קרן"].astype(str).str.zfill(8); f = f.set_index("מס_קרן")
    for c in ("שם", "מנהל", "סיווג", "שווי_נכסים_אלפי_ש״ח", "חשיפה_מניות_מדווחת", "חשיפה_מטח_מדווחת", "מחמ_אגח_מדווח", "תאריך_דוח"):
        if c in f.columns: out[c] = f[c].reindex(out.index)
    out.index.name = "מס_קרן"
    return out.reset_index()


def main():
    data = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "data"); fdir = data / "funds"
    allp = []
    for mdir in sorted(p for p in fdir.iterdir() if p.is_dir() and len(p.name) == 7 and p.name[4] == "-"):
        cp = mdir / "composition.csv"
        if not cp.exists() or "--force" in sys.argv:
            hp, fp = mdir / "holdings.csv.gz", mdir / "funds.csv"
            if not hp.exists() or not fp.exists(): continue
            h = pd.read_csv(hp, dtype=str, keep_default_na=False); h["מס_קרן"] = h["מס_קרן"].str.zfill(8)
            fdf = pd.read_csv(fp, dtype=str, encoding="utf-8-sig")
            c = composition(h, fdf); c.insert(0, "חודש", mdir.name); c.to_csv(cp, index=False, encoding="utf-8-sig")
            print(mdir.name, len(c), "קרנות", flush=True)
        allp.append(pd.read_csv(cp, dtype={"מס_קרן": str}, encoding="utf-8-sig"))
    if allp:
        a = pd.concat(allp, ignore_index=True); a["מס_קרן"] = a["מס_קרן"].str.zfill(8)
        a.to_csv(fdir / "composition_all.csv", index=False, encoding="utf-8-sig")
        print("composition_all:", len(a), "שורות,", a["חודש"].nunique(), "חודשים")


if __name__ == "__main__":
    main()
