"""Turn a maya-fetch output folder into the CORP issuer-folder layout.

    python maya/sort_to_corp.py <fetch_dir/<companyId>_<name>> <staging_dir> [--existing hashes.txt]

Writes <staging_dir>/דוחות/<sub>/<YYYY-MM-DD>_<type>_<desc>.<ext> and <staging_dir>/maya_manifest.csv.
--existing: optional file of "md5  name" lines (files already in the issuer folder) -> duplicates are skipped.
Naming follows the CORP convention: periodic reports are dated by period end (2026-06-30_רבעוני_Q2-2026.pdf),
everything else by publish date.
"""
import csv, hashlib, pathlib, re, shutil, sys

SUB = {"שנתי": "שנתי", "רבעוני": "רבעוני", "מיידי": "מיידי", "אחר": "אחר", "דירוג": "דירוג", "שווי": "שווי_והערכות"}
QEND = {1: "03-31", 2: "06-30", 3: "09-30", 4: "12-31"}

def short(s, n=55):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s or "")
    s = s.replace("״", "").replace('"', "")
    s = re.sub(r"\s*[-–]\s*", "-", s.strip())
    s = re.sub(r"[\s,]+", "-", s)
    s = re.sub(r"-{2,}", "-", s)
    return s[:n].strip("-.")

def asof(title):
    m = re.search(r"ליום\s*(\d{1,2}\.\d{1,2}\.\d{2,4})", title)
    return m.group(1) if m else ""

def classify(row):
    """-> (sub, type_tag, date, desc)"""
    form, title, pub = row["form"], row["title"], row["date"][:10]
    if form == "ת930":
        y = re.search(r"(20\d\d)", title); y = y.group(1) if y else pub[:4]
        if re.search(r"תקופתי|שנתי לשנת", title) and "חצי" not in title:
            return "שנתי", "שנתי", f"{y}-12-31", f"דוח-תקופתי-{y}"
        q = re.search(r"רבעון\s*(\d)", title); q = int(q.group(1)) if q else 0
        if q in QEND:
            return "רבעוני", "רבעוני", f"{y}-{QEND[q]}", f"Q{q}-{y}"
        return "רבעוני", "רבעוני", pub, short(title)
    if form == "ת126":
        return "מיידי", "מיידי", pub, "מצבת-התחייבויות-" + (asof(title) or short(title, 30))
    if form == "ת077":
        return "מיידי", "מיידי", pub, "מצבת-בעלי-עניין-" + (asof(title) or short(title, 30)) + ("-תיקון" if "תיקון" in title else "")
    if form in ("ת125", "ד125") or row["reason"] == "דירוג":
        agency = "מידרוג" if "מידרוג" in title else ("מעלות" if re.search("מעלות|S&P", title) else "")
        return "דירוג", "דירוג", pub, short((agency + "-" if agency and agency not in title[:12] else "") + title, 60) + ("-דוח-החברה" if form == "ת125" else "")
    if row["reason"] == "מצגת":
        return "אחר", "אחר", pub, short(title, 50)
    if row["reason"] == "שומה":
        return "שווי", "שווי", pub, short(title, 55)
    if form.startswith("א"):
        return "אחר" if "נאמן" in title and "דו" in title else "מיידי", "אחר" if "נאמן" in title and "דו" in title else "מיידי", pub, short(title, 55)
    return "מיידי", "מיידי", pub, short(title, 55)

def att_suffix(kind, ext, n_pdf, idx_pdf, fname):
    if ext == "xbrl":
        return ""
    if ext == "htm":
        return "_שער-מגנא" if n_pdf else ""
    if ext == "txt":
        return "_טקסט"
    if ext == "pdf" and n_pdf > 1 and idx_pdf > 0:
        hint = short(re.sub(r"_isa|\.pdf$", "", fname or "", flags=re.I), 30)
        return f"_נספח-{idx_pdf}" + (f"-{hint}" if hint else "")
    return ""

def main():
    src, stg = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
    existing = {}
    if "--existing" in sys.argv:
        for line in open(sys.argv[sys.argv.index("--existing") + 1], encoding="utf-8"):
            if line.strip():
                h, n = line.split(None, 1); existing[h] = n.strip()
    rows = list(csv.DictReader(open(src / "catalog.csv", encoding="utf-8-sig")))
    man, used, seen_h = [], set(), {}
    for r in rows:
        if r["keep"] != "כן" or not r["saved"]:
            continue
        sub, tag, date, desc = classify(r)
        atts = [a for a in r["attachments"].split(" | ") if a]
        saved = r["saved"].split(" | ")
        kinds = [a.split(":")[0] for a in atts]
        fnames = [a.split(":")[1] if a.count(":") >= 2 else "" for a in atts]
        n_pdf = sum(1 for k in kinds if k.startswith("pdf")); i_pdf = 0
        for k, (kind, fname, sv) in enumerate(zip(kinds, fnames, saved)):
            if sv.startswith(("FAILED", "TOO_BIG")):
                man.append([r["report_id"], r["date"][:10], r["form"], r["title"], "", sv, r["maya_url"]]); continue
            p = src / sv; ext = p.suffix.lstrip(".").lower()
            suf = att_suffix(kind, ext, n_pdf, i_pdf, fname)
            if ext == "pdf": i_pdf += 1
            name = f"{date}_{tag}_{desc}{suf}.{ext}"
            base = name; j = 2
            while (sub, name) in used:
                name = base.replace(f".{ext}", f"_{j}.{ext}"); j += 1
            h = hashlib.md5(p.read_bytes()).hexdigest()
            if h in existing:
                status = "כפול-קיים:" + existing[h]
            elif h in seen_h:
                status = "כפול-בריצה:" + seen_h[h]
            else:
                used.add((sub, name)); seen_h[h] = name
                dst = stg / "דוחות" / SUB[sub] / name; dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, dst); status = f"דוחות\\{SUB[sub]}\\{name}"
            man.append([r["report_id"], r["date"][:10], r["form"], r["title"], kind, status, r["maya_url"]])
    with open(stg / "maya_manifest.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh); w.writerow(["report_id", "publish_date", "form", "title", "attachment", "saved_as", "maya_url"]); w.writerows(man)
    print(f"{sum(1 for m in man if m[5].startswith('דוחות'))} files staged, "
          f"{sum(1 for m in man if m[5].startswith('כפול'))} duplicates skipped")

if __name__ == "__main__":
    main()
