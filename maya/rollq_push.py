"""Keep a copy of one periodic report's files in corp-data until the model update (roll) succeeds.

    python maya/rollq_push.py <src_dir> <issuer_id> <period> <remote_url> [--dry]

The PC deletes deliver/* branches after copying them, so the roll task needs its own copy: one orphan commit on
rollq/<issuer_id>-<period> holding the files flat. The roll task deletes the branch after a successful update.
Cover pages (*_שער-מגנא.htm) and text dumps (*_טקסט.txt) are skipped, as in the daily run. Prints the branch name.
"""
import argparse, pathlib, shutil, sys, tempfile
from push_deliver import git

SKIP = ("_שער-מגנא.htm", "_טקסט.txt")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("issuer_id"); ap.add_argument("period"); ap.add_argument("remote")
    ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()
    files = sorted(p for p in pathlib.Path(a.src).rglob("*") if p.is_file() and not p.name.endswith(SKIP))
    branch = f"rollq/{a.issuer_id}-{a.period}"
    print(branch, len(files), "files")
    for p in files: print(" ", p.name)
    if a.dry or not files: return 0
    with tempfile.TemporaryDirectory() as tmp:
        git(tmp, "init", "-q", "-b", "rollq")
        for p in files: shutil.copy2(p, pathlib.Path(tmp) / p.name)
        git(tmp, "add", "-A"); git(tmp, "commit", "-q", "-m", f"rollq {a.issuer_id} {a.period}: {len(files)} files")
        git(tmp, "push", "-q", "-f", a.remote, f"HEAD:refs/heads/{branch}")
    return 0


if __name__ == "__main__":
    sys.path.insert(0, str(pathlib.Path(__file__).parent))
    sys.exit(main())
