import json, sys, pathlib, requests
OUT = pathlib.Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")
B = "https://maya.tase.co.il"
S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept": "application/json", "Referer": B + "/he/companies", "Origin": B})
res = {}
for q in ["נמקו", "נמקו ריאלטי", "namco", "nemco"]:
    try:
        r = S.get(B + "/api/v1/companies/autocomplete", params={"search": q, "take": 10}, timeout=60)
        res[q] = dict(status=r.status_code, body=r.text[:4000])
    except Exception as e:
        res[q] = dict(err=str(e)[:300])
(OUT / "find.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print(json.dumps(res, ensure_ascii=False)[:3000])
