import pathlib, subprocess, sys
ROOT = pathlib.Path(__file__).resolve().parents[1]

def test_dry_lists_kept_files(tmp_path):
    for n in ("2026-09-30_רבעוני_Q3-2026.pdf", "2026-09-30_רבעוני_Q3-2026_טופס.htm", "2026-09-30_רבעוני_Q3-2026.xbrl",
              "2026-11-11_מיידי_ת126.htm", "x_שער-מגנא.htm", "y_טקסט.txt"):
        (tmp_path / n).write_text("x", encoding="utf-8")
    out = subprocess.run([sys.executable, str(ROOT / "maya" / "rollq_push.py"), str(tmp_path), "513821488", "2026Q3", "file:///dev/null", "--dry"],
                         capture_output=True, text=True, check=True).stdout
    assert "rollq/513821488-2026Q3" in out
    assert "Q3-2026.pdf" in out and "ת126" in out
    assert "שער-מגנא" not in out and "טקסט.txt" not in out
