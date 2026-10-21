"""yahoo_prices.py — closes for given Yahoo symbols on given dates (for CORP holdings valuation). usage: python maya/yahoo_prices.py <out_dir>"""
import subprocess, sys, os, json
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "yfinance"], check=False)
import yfinance as yf
out = sys.argv[1] if len(sys.argv) > 1 else "probe_out"
os.makedirs(out, exist_ok=True)
SYMS = ["NWMD.TA", "ITH.L", "ISCD.TA", "MEDN.TA", "MEDN-M.TA", "IPO.TO", "IPOOF", "DLEKG.TA", "DLKIS.TA", "USDILS=X", "GBPILS=X", "CADILS=X", "GBPUSD=X"]
DATES = ["2025-12-31", "2026-06-30", "2026-10-20"]
res = {}
for s in SYMS:
    try:
        h = yf.Ticker(s).history(start="2025-12-20", end="2026-10-22", auto_adjust=False)
        if h.empty: res[s] = {"error": "empty"}; continue
        h.index = h.index.tz_localize(None) if h.index.tz is not None else h.index
        r = {}
        for d in DATES:
            sub = h[h.index <= d]
            if len(sub): r[d] = {"date": str(sub.index[-1].date()), "close": float(sub["Close"].iloc[-1])}
        info = {}
        try:
            fi = yf.Ticker(s).fast_info
            info = {"currency": fi.get("currency"), "shares": fi.get("shares"), "mcap": fi.get("marketCap")}
        except Exception as e: info = {"err": str(e)}
        res[s] = {"px": r, "info": info, "last": str(h.index[-1].date())}
    except Exception as e:
        res[s] = {"error": str(e)}
json.dump(res, open(os.path.join(out, "yahoo_prices.json"), "w"), indent=1, default=str)
print(json.dumps(res, indent=1, default=str))
