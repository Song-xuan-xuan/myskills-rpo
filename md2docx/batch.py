#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
batch.py —— 一份 md，同时出多个版本的 Word

    python3 batch.py 报告.md                                  # 三套内置档案全出
    python3 batch.py 报告.md --profiles technical,academic     # 指定档案
    python3 batch.py 报告.md --outdir dist --prefix 2026Q3     # 指定输出目录和前缀
    python3 batch.py 报告.md --profiles a.json,b.json --sample # 每个档案再出一张样张

输出（默认）：
    dist/报告_technical.docx
    dist/报告_official.docx
    dist/报告_academic.docx

每个版本用同一份 md，差别只在档案（字体/编号/封面/页眉）。
适合「给领导一版、给甲方一版、投稿一版」这种场景。
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "profiles")


def die(msg):
    print(msg, file=sys.stderr)
    sys.exit(1)


def builtin_profiles():
    out = []
    if os.path.isdir(PROFILE_DIR):
        for f in sorted(os.listdir(PROFILE_DIR)):
            if f.endswith(".json") and not f.startswith("_"):
                out.append(os.path.splitext(f)[0])
    return out


def resolve(p):
    """档案名 / 路径都行"""
    for c in (p, os.path.join(PROFILE_DIR, p), os.path.join(PROFILE_DIR, p + ".json"),
              os.path.join(os.getcwd(), p), os.path.join(os.getcwd(), p + ".json")):
        if os.path.isfile(c):
            return os.path.abspath(c)
    return None


def main():
    argv = sys.argv[1:]
    if not argv or argv[0].startswith("-"):
        print(__doc__)
        return 1

    md = os.path.abspath(argv[0])
    if not os.path.isfile(md):
        die("找不到 md：%s" % md)

    profiles = builtin_profiles()
    if "--profiles" in argv:
        profiles = [x.strip() for x in argv[argv.index("--profiles") + 1].split(",")
                    if x.strip()]
    outdir = "dist"
    if "--outdir" in argv:
        outdir = argv[argv.index("--outdir") + 1]
    prefix = os.path.splitext(os.path.basename(md))[0]
    if "--prefix" in argv:
        prefix = argv[argv.index("--prefix") + 1]
    do_sample = "--sample" in argv
    extra = []
    if "--" in argv:
        extra = argv[argv.index("--") + 1:]

    outdir = os.path.abspath(outdir)
    os.makedirs(outdir, exist_ok=True)

    print("md：%s" % md)
    print("出 %d 个版本 → %s" % (len(profiles), outdir))
    print()

    made, failed = [], []
    for name in profiles:
        path = resolve(name)
        if not path:
            failed.append((name, "找不到档案"))
            continue
        label = os.path.splitext(os.path.basename(path))[0]
        out = os.path.join(outdir, "%s_%s.docx" % (prefix, label))
        cmd = [sys.executable, os.path.join(HERE, "md2docx.py"), md, out,
               "--profile", path] + extra
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0 or not os.path.exists(out):
            failed.append((label, (r.stderr or r.stdout).strip()[-200:]))
            print("  ✗ %-12s %s" % (label, (r.stderr or r.stdout).strip()[:80]))
            continue
        size = os.path.getsize(out) // 1024
        made.append(out)
        print("  ✓ %-12s %s（%d KB）" % (label, os.path.basename(out), size))

    if do_sample:
        print("\n顺带出样张：")
        for name in profiles:
            path = resolve(name)
            if not path:
                continue
            r = subprocess.run([sys.executable, os.path.join(HERE, "configure.py"),
                                "--sample", path], capture_output=True, text=True)
            if r.returncode == 0:
                print("  ✓ 样张_%s" % os.path.splitext(os.path.basename(path))[0])

    print()
    if made:
        print("完成 %d 个版本：" % len(made))
        for m in made:
            print("  %s" % m)
    if failed:
        print("\n失败 %d 个：" % len(failed))
        for name, why in failed:
            print("  %s：%s" % (name, why))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
