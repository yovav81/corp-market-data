"""Push sorted report files to deliver/* branches of corp-data; the PC sync copies them into the issuers'
דוחות folders and deletes the branches.

    python maya/push_deliver.py <staging_root> <remote_url> <name> [--max-mb 40]

<staging_root> holds the final tree, e.g. 02_Issuers/<corp_folder>/דוחות/<sub>/<file>.
Each branch is one orphan commit; files are split into branches of at most --max-mb (one file may exceed it).
Prints the branch names that were pushed.
"""
import argparse, pathlib, shutil, subprocess, sys, tempfile


def git(cwd, *args):
    subprocess.run(["git", "-c", "user.name=CORP bot", "-c", "user.email=corp-bot@users.noreply.github.com",
                    "-c", "commit.gpgsign=false", "-c", "core.quotepath=false", *args], cwd=cwd, check=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root"); ap.add_argument("remote"); ap.add_argument("name")
    ap.add_argument("--max-mb", type=float, default=40)
    a = ap.parse_args()
    root = pathlib.Path(a.root)
    files = sorted(p for p in root.rglob("*") if p.is_file())
    if not files:
        print("nothing to deliver"); return 0
    bad = [p for p in files if not (len(p.relative_to(root).parts) >= 4 and p.relative_to(root).parts[0] == "02_Issuers" and p.relative_to(root).parts[2] == "דוחות")]
    if bad:
        print(f"refusing: {len(bad)} files outside 02_Issuers/<issuer>/דוחות/ — e.g. {bad[0].relative_to(root)}", file=sys.stderr); return 2
    limit = a.max_mb * 1024 * 1024
    chunks, cur, size = [], [], 0
    for p in files:
        s = p.stat().st_size
        if cur and size + s > limit:
            chunks.append(cur); cur, size = [], 0
        cur.append(p); size += s
    if cur:
        chunks.append(cur)
    for k, chunk in enumerate(chunks, 1):
        branch = f"deliver/{a.name}" + (f"-{k}" if len(chunks) > 1 else "")
        with tempfile.TemporaryDirectory() as tmp:
            git(tmp, "init", "-q", "-b", "deliver")
            for p in chunk:
                dst = pathlib.Path(tmp) / p.relative_to(root)
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, dst)
            git(tmp, "add", "-A")
            git(tmp, "commit", "-q", "-m", f"deliver {a.name}: {len(chunk)} files")
            git(tmp, "push", "-q", a.remote, f"HEAD:refs/heads/{branch}")
        print(branch, len(chunk), "files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
