"""
short_weekly.py — הטבלה השבועית (יום ראשון) לטלגרם: כל האג"ח הקונצרניות עם מח"מ מתחת ל-1, כתמונה (כמו טבלת המרווחים החריגים) + אקסל.

קורא data/short/short_bonds.csv (מהריצה האחרונה) ו-data/hist/spread_stats.csv (דירוג מנורמל, מסוקר), וכותב:
  data/short/weekly_table*.png   הטבלה כתמונה (עד 25 שורות לתמונה), ממוינת לפי מרווח שקלי צפוי
  data/short/weekly_table.xlsx   אותה טבלה
  data/short/weekly_message.txt  הכיתוב לטלגרם
הרצה: python short/short_weekly.py data
"""
from __future__ import annotations
import datetime as dt, sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "hist"))
from alert_render import render_png, ROWS_PER_IMG  # noqa: E402
from ratings import lower_of  # noqa: E402

COLS = [("ני\"ע", "sid", 95, "c"), ("סדרה", "name", 250, "r"), ("דירוג", "rating", 80, "c"), ("הצמדה", "link", 75, "c"), ("מח\"מ", "dur", 70, "c"),
        ("ימים", "days", 65, "c"), ("תשואה %", "ytm", 95, "c"), ("שקלי צפוי %", "ytm_exp", 125, "c"), ("רווח לפדיון %", "tot", 130, "c"),
        ("מרווח נ\"ב", "spr", 110, "c"), ("±0.1% מדד", "sens", 115, "c"), ("תשלום", "nxt", 90, "c"), ("% קרן", "pct", 70, "c"), ("★ / הערה", "flag", 175, "c")]
LINK = {"מדד": "צמוד", "שקלי": "שקלי", "דולר": "מט\"ח"}


def fmt(v, f="{:.2f}"):
    return "—" if v is None or (isinstance(v, float) and v != v) else f.format(v)


def main(data="data"):
    d = Path(data); sd = d / "short"
    sb = pd.read_csv(sd / "short_bonds.csv", encoding="utf-8-sig")
    try:
        hs = pd.read_csv(d / "hist" / "spread_stats.csv", encoding="utf-8-sig")[["security_id", "rating_low", "covered"]]
        hs["security_id"] = pd.to_numeric(hs["security_id"], errors="coerce"); hs = hs.drop_duplicates("security_id")
        sb = sb.merge(hs, on="security_id", how="left")
    except Exception:
        sb["rating_low"] = None; sb["covered"] = ""
    date = str(sb["date"].iloc[0]); dd = dt.date.fromisoformat(date).strftime("%d.%m.%Y")
    ok = sb[sb["status"] == "ok"].copy()
    ok["flags"] = ok["flags"].fillna("")
    dist = ok[ok["flags"].str.contains("מצוקה")]; ok = ok[~ok["flags"].str.contains("מצוקה")]
    ok["sort"] = ok["spread_nominal_bp"].fillna(ok["spread_real_bp"]).fillna(-999)
    ok = ok.sort_values("sort", ascending=False)
    rows = []
    for _, r in pd.concat([ok, dist]).iterrows():
        rating = r.get("rating_low")
        if rating is None or (isinstance(rating, float) and rating != rating):
            rating = lower_of(r.get("rating_maalot"), r.get("rating_midroog")) or "—"
        note = []
        if "מצוקה" in r["flags"]: note.append("מצוקה")
        if "משתנה" in r["flags"]: note.append("ר' משתנה")
        if "מט״ח" in r["flags"]: note.append("מט\"ח")
        if "פער מול גלובס" in r["flags"]: note.append("פער")
        if "10 ימים" in r["flags"]: note.append("תשלום בקרוב")
        if "טווח" in r["flags"]: note.append("מדד אחד")
        nxt = r.get("next_pay_date"); nxt = dt.date.fromisoformat(nxt).strftime("%d.%m") if isinstance(nxt, str) else "—"
        rows.append({"sid": str(int(r["security_id"])), "name": r["name"], "rating": rating, "link": LINK.get(str(r.get("linkage")), str(r.get("linkage"))),
                     "dur": fmt(r.get("dur_calc"), "{:.2f}"), "ytm": fmt(r.get("ytm_globes")), "ytm_exp": fmt(r.get("ytm_nominal_exp")),
                     "spr": fmt(r.get("spread_nominal_bp") if r.get("spread_nominal_bp") == r.get("spread_nominal_bp") else r.get("spread_real_bp"), "{:+.0f}"),
                     "days": fmt(r.get("days_to_final"), "{:.0f}"), "tot": fmt(r.get("total_return_pct")),
                     "sens": (fmt(r.get("cpi_sens_bp"), "±{:.0f}") if "מדד" in str(r.get("linkage")) else "—"), "nxt": nxt, "pct": fmt(r.get("next_principal_pct"), "{:.0f}"),
                     "covered": ("★" if isinstance(r.get("covered"), str) and r["covered"] else ""), "note": " · ".join(note),
                     "maya_recent": ""})
    for x in rows:
        x["flag_text"] = (x["covered"] + " " + x["note"]).strip()
    sd.mkdir(exist_ok=True)
    for old in sd.glob("weekly_table*.png"): old.unlink()
    import alert_render
    orig = alert_render.flag_of
    alert_render.flag_of = lambda r: r.get("flag_text", "")
    try:
        sub = ("תשואה לפדיון = גלובס (צמוד: ריאלית); שקלי צפוי = אותו תזרים עם המדדים הצפויים (בנק ישראל + עונתיות), שנתי; רווח לפדיון = סה\"כ עד התשלום האחרון, לא שנתי | "
               "מרווח = מול ממשלתי שקלי באותו מח\"מ, נ\"ב; ±0.1% מדד = כמה נ\"ב זז השקלי הצפוי לכל 0.1% סטייה במדד. ׳מדד אחד׳ = פחות מ-60 יום. ממוין לפי מרווח; מצוקה בסוף. ★ מסוקר")
        pngs = render_png(rows, f"אג\"ח קונצרניות עם מח\"מ מתחת ל-1 — {dd}", sub, str(sd / "weekly_table"), cols=COLS)
    finally:
        alert_render.flag_of = orig
    # אקסל
    cols = ["security_id", "name", "rating_low", "linkage", "redemption_type", "coupon", "last_price", "dur_globes", "ytm_globes", "dur_calc", "ytm_real_calc",
            "ytm_nominal_exp", "days_to_final", "total_return_pct", "cpi_sens_bp", "infl_annual_pct", "gov_ils_at_dur", "spread_nominal_bp", "spread_real_bp", "next_pay_date", "next_principal_pct", "next_ex",
            "final_pay_date", "gap_bp", "covered", "flags", "volume_k_nis"]
    heb = ["מס׳ ני״ע", "סדרה", "דירוג", "הצמדה", "סוג פדיון", "ריבית %", "מחיר", "מח״מ גלובס", "תשואה גלובס %", "מח״מ מחושב", "תשואה ריאלית מחושבת %",
           "תשואה שקלית צפויה %", "ימים לפדיון", "רווח כולל לפדיון % (לא שנתי)", "רגישות +0.1% מדד (נ״ב)", "אינפלציה גלומה שנתית %", "ממשלתי שקלי באותו מח״מ %", "מרווח שקלי צפוי (נ״ב)", "מרווח ריאלי (נ״ב)", "תשלום קרוב", "% קרן בתשלום",
           "אקס קרוב", "פדיון סופי", "פער מול גלובס (נ״ב)", "מסוקר", "דגלים", "מחזור (אלפי ₪)"]
    out = pd.concat([ok, dist])[cols].copy(); out.columns = heb
    with pd.ExcelWriter(sd / "weekly_table.xlsx", engine="openpyxl") as xw:
        out.to_excel(xw, index=False, sheet_name="מחמ מתחת ל 1")
        ws = xw.sheets["מחמ מתחת ל 1"]; ws.sheet_view.rightToLeft = True; ws.freeze_panes = "C2"
        for i, w in enumerate([11, 24, 8, 8, 14, 8, 8, 9, 10, 9, 12, 12, 8, 12, 11, 12, 12, 12, 11, 11, 9, 11, 11, 10, 7, 50, 10]):
            ws.column_dimensions[ws.cell(1, i + 1).column_letter].width = w
    top = ok.head(5)
    lines = "\n".join(f"• {r['name']}: {fmt(r['ytm_nominal_exp'])}% שקלי צפוי ({fmt(r['total_return_pct'])}% ב-{fmt(r['days_to_final'], '{:.0f}')} ימים), {fmt(r['sort'], '{:+.0f}')} נ\"ב" for _, r in top.iterrows())
    msg = (f"📅 <b>אג\"ח קצרות (מח\"מ עד שנה) — {dd}</b>  ({len(ok)} סדרות + {len(dist)} במצוקה)\n"
           f"<i>תשואה שקלית צפויה עם המדדים הצפויים; מרווח מול ממשלתי שקלי באותו מח\"מ. המובילות:</i>\n{lines}")
    (sd / "weekly_message.txt").write_text(msg, encoding="utf-8")
    print(f"{len(rows)} שורות, {len(pngs)} תמונות")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data")
