"""maya/watch.py: מפתח מנפיק שלא קיים ב-issuers.json חייב להכשיל את הריצה — לא לעבור בשקט.

רקע (11.11.2026): ריצת האתחול של אקסטל (init=true, only=extell) רצה על קומיט שבו extell עוד לא היה
ב-issuers.json. הלולאה דילגה על כל המנפיקים, הריצה הסתיימה כהצלחה, ולא נרשם קו בסיס ב-state.json.
"""
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAYA = ROOT / "maya"
sys.path.insert(0, str(MAYA))

ISSUERS = [{"key": "big"}, {"key": "extell"}]


def run_watch(tmp_path, only):
    """מריץ את הסקריפט האמיתי; מחזיר (תהליך, תיקיית הפלט, תוכן קובצי המצב לפני ואחרי)."""
    out = tmp_path / "maya_out"
    files = [MAYA / "state.json", MAYA / "issuers.json"]
    before = [f.read_bytes() for f in files]
    proc = subprocess.run([sys.executable, str(MAYA / "watch.py"), "--out", str(out), "--init", "--only", only],
                          capture_output=True, text=True, timeout=180, env={**os.environ, "GITHUB_ACTIONS": "true"})
    after = [f.read_bytes() for f in files]
    return proc, out, before, after


def test_unknown_only_key_fails_before_any_work(tmp_path):
    proc, out, before, after = run_watch(tmp_path, "nosuchissuer")
    text = proc.stdout + proc.stderr
    assert proc.returncode != 0
    assert "nosuchissuer" in text                      # ההודעה אומרת איזה מפתח לא נמצא
    assert any(line.startswith("::error::") and "nosuchissuer" in line for line in text.splitlines())   # מופיע בסיכום הריצה ב-GitHub
    assert not out.exists()                            # אין תיקיית פלט -> ה-workflow לא דוחף ענף ריק
    assert before == after                             # state.json / issuers.json לא נגעו


def test_one_unknown_key_among_known_ones_fails(tmp_path):
    known = json.loads((MAYA / "issuers.json").read_text(encoding="utf-8"))[0]["key"]
    proc, out, before, after = run_watch(tmp_path, f"{known},nosuchissuer")
    assert proc.returncode != 0
    assert "nosuchissuer" in proc.stdout + proc.stderr
    assert not out.exists()
    assert before == after


def test_known_keys_and_no_filter_are_accepted():
    from watch import unknown_keys
    assert unknown_keys(ISSUERS, "") == []             # הריצה היומית: בלי סינון = כל המנפיקים
    assert unknown_keys(ISSUERS, "extell") == []
    assert unknown_keys(ISSUERS, "big,extell") == []
    assert unknown_keys(ISSUERS, "extell,") == []      # פסיק עודף אינו מפתח


def test_unknown_keys_are_reported_by_name():
    from watch import unknown_keys
    assert unknown_keys(ISSUERS, "big,nope,extell,typo") == ["nope", "typo"]
