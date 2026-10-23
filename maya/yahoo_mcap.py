"""yahoo_mcap.py — market cap / shares / close from Yahoo for CORP issuers. usage: python maya/yahoo_mcap.py <out_dir>"""
import subprocess, sys, os, json
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "yfinance"], check=False)
import yfinance as yf
out = sys.argv[1] if len(sys.argv) > 1 else "probe_out"
os.makedirs(out, exist_ok=True)
SYMS = os.environ.get("SYMS", "TDRN.TA,TIDR.TA,TDHR.TA,TDR.TA").split(",")
res = {}
for s in SYMS:
    try:
        t = yf.Ticker(s)
        h = t.history(period="1mo", auto_adjust=False)
        r = {"last": [[str(i.date()), float(c)] for i, c in zip(h.index[-5:], h["Close"].iloc[-5:])] if len(h) else None}
        try:
            fi = t.fast_info; r["fast"] = {"currency": fi.get("currency"), "shares": fi.get("shares"), "mcap": fi.get("marketCap"), "last_price": fi.get("lastPrice")}
        except Exception as e: r["fast_err"] = str(e)
        try:
            inf = t.info; r["info"] = {k: inf.get(k) for k in ("longName", "shortName", "marketCap", "sharesOutstanding", "currency", "regularMarketPrice", "previousClose")}
        except Exception as e: r["info_err"] = str(e)
        res[s] = r
    except Exception as e:
        res[s] = {"error": str(e)}
json.dump(res, open(os.path.join(out, "yahoo_mcap.json"), "w"), indent=1, default=str)
print(json.dumps(res, indent=1, default=str))
