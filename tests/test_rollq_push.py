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


def test_existing_branch_keeps_files(tmp_path):
    remote = tmp_path / "r.git"; subprocess.run(["git", "init", "-q", "--bare", str(remote)], check=True)
    a = tmp_path / "a"; a.mkdir(); (a / "report.pdf").write_text("r", encoding="utf-8")
    b = tmp_path / "b"; b.mkdir(); (b / "fix.pdf").write_text("f", encoding="utf-8")
    exe = str(ROOT / "maya" / "rollq_push.py")
    subprocess.run([sys.executable, exe, str(a), "1", "2026Q3", str(remote)], check=True, capture_output=True)
    subprocess.run([sys.executable, exe, str(b), "1", "2026Q3", str(remote)], check=True, capture_output=True)
    files = subprocess.run(["git", "--git-dir", str(remote), "ls-tree", "--name-only", "rollq/1-2026Q3"], capture_output=True, text=True, check=True).stdout.split()
    assert sorted(files) == ["fix.pdf", "report.pdf"]
