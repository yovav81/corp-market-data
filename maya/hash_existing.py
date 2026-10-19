"""md5 lines for sort_to_corp.py --existing: "md5  name" for every file, plus "t:<md5 of visible text>  name"
for .htm files (the same Maya form saved twice differs in markup bytes but not in text).

    python maya/hash_existing.py <file> [<file> ...] > hashes.txt
"""
import hashlib, html, re, sys

def text_hash(b):
    for enc in ("utf-8", "cp1255"):
        try:
            t = b.decode(enc); break
        except UnicodeDecodeError:
            continue
    else:
        t = b.decode("latin1")
    t = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", html.unescape(t)).strip()
    return hashlib.md5(t.encode("utf-8")).hexdigest()

if __name__ == "__main__":
    for f in sys.argv[1:]:
        b = open(f, "rb").read(); name = f.replace("\\", "/").rsplit("/", 1)[-1]
        print(hashlib.md5(b).hexdigest(), name)
        if f.lower().endswith((".htm", ".html")):
            print("t:" + text_hash(b), name)
