#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make.py —— 一条命令从 md 到「可交付的 docx」

    python3 make.py 报告.md                          # 转换 + 质检
    python3 make.py 报告.md --profile 客户A.json      # 指定档案
    python3 make.py 报告.md --pdf                    # 顺带导 PDF（需要本机 Word）
    python3 make.py 报告.md --sheet                  # 再加一张关键页联系表
    python3 make.py 报告.md --profiles a,b,c          # 一份 md 出多版
    python3 make.py 报告.md --quiet                   # 只输出结论

做的事：
  1. 转换（md2docx.py）
  2. 结构合规检查（跨部件引用、元素顺序、书签、REF 目标…）
  3. 交付质量检查（目录制表位、图片是否超版心、页脚是否有页码域…）
  4. 打印人话摘要：本次用了哪套格式、成品在哪、打开后要做什么

任何一步失败都会明确报出来，并以非 0 退出。
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


def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def resolve_profile(name):
    if not name:
        return None
    for c in (name, os.path.join(PROFILE_DIR, name),
              os.path.join(PROFILE_DIR, name + ".json"),
              os.path.join(os.getcwd(), name),
              os.path.join(os.getcwd(), name + ".json")):
        if os.path.isfile(c):
            return os.path.abspath(c)
    die("找不到档案：%s" % name)


def main():
    argv = sys.argv[1:]
    if not argv or argv[0].startswith("-"):
        print(__doc__)
        return 1
    md = os.path.abspath(argv[0])
    if not os.path.isfile(md):
        die("找不到 md：%s" % md)

    profile = None
    profiles = None
    out = None
    do_pdf = "--pdf" in argv or "--sheet" in argv
    quiet = "--quiet" in argv
    outdir = None
    for flag, key in (("--profile", "profile"), ("--profiles", "profiles"),
                      ("--out", "out"), ("--outdir", "outdir")):
        if flag in argv:
            i = argv.index(flag)
            if i + 1 >= len(argv):
                die("%s 后面要跟参数" % flag)
            val = argv[i + 1]
            if key == "profile":
                profile = resolve_profile(val)
            elif key == "profiles":
                profiles = [x.strip() for x in val.split(",") if x.strip()]
            elif key == "out":
                out = os.path.abspath(val)
            else:
                outdir = os.path.abspath(val)

    # ---------- 1. 转换 ----------
    made = []
    if profiles:
        base = os.path.splitext(os.path.basename(md))[0]
        d = outdir or os.path.dirname(md)
        cmd = [sys.executable, os.path.join(HERE, "batch.py"), md,
               "--profiles", ",".join(profiles), "--outdir", d]
        r = run(cmd)
        if r.returncode != 0:
            print(r.stdout[-800:] + r.stderr[-800:], file=sys.stderr)
            die("批量转换失败")
        if not quiet:
            print(r.stdout.rstrip())
        for p in profiles:
            path = resolve_profile(p)
            label = os.path.splitext(os.path.basename(path))[0]
            fp = os.path.join(d, "%s_%s.docx" % (base, label))
            if os.path.exists(fp):
                made.append((fp, path))
    else:
        if not out:
            out = os.path.splitext(md)[0] + ".docx"
        # 显式指定档案：这样不依赖任何侧车文件也能复现
        if not profile:
            profile = resolve_profile("technical")
        cmd = [sys.executable, os.path.join(HERE, "md2docx.py"), md, out,
               "--profile", profile]
        r = run(cmd)
        if r.returncode != 0:
            print(r.stderr[-1200:] or r.stdout[-1200:], file=sys.stderr)
            die("转换失败")
        if r.stderr.strip() and not quiet:
            print(r.stderr.strip(), file=sys.stderr)
        made.append((out, profile))

    # ---------- 2+3. 质检 ----------
    ok = True
    report = []
    for path, prof in made:
        r = run([sys.executable, os.path.join(HERE, "verify.py"), path])
        good = r.returncode == 0
        ok = ok and good
        report.append((path, prof, good, r.stdout))

    # ---------- 4. 摘要 ----------
    print()
    for path, prof, good, detail in report:
        size = os.path.getsize(path) // 1024 if os.path.exists(path) else 0
        print("%s %s（%d KB）" % ("✓" if good else "✗", path, size))
        if prof:
            r = run([sys.executable, os.path.join(HERE, "format_report.py"), prof])
            lines = [l for l in r.stdout.splitlines()
                     if l.startswith(("正文", "标题1", "图题注", "表格　", "页眉",
                                      "目录　", "封面　", "参考文献"))]
            if lines:
                print("    本次格式：")
                for l in lines[:8]:
                    print("      " + l.strip())
        if not good:
            for l in detail.splitlines():
                if l.strip().startswith("✗"):
                    print("    " + l.strip())

    # ---------- 可选 PDF ----------
    if do_pdf:
        print()
        for path, _ in made:
            # 用户明确要 PDF/联系表，才允许驱动 Word
            extra = ["--word"]
            extra += ["--sheet", "--no-open"] if "--sheet" in argv else ["--no-open"]
            r = run([sys.executable, os.path.join(HERE, "preview.py"), path] + extra)
            pdf = os.path.splitext(path)[0] + ".pdf"
            sheet = os.path.splitext(path)[0] + "_联系表.png"
            if r.returncode == 0 and os.path.exists(pdf):
                print("✓ PDF：%s" % pdf)
                if os.path.exists(sheet):
                    print("✓ 联系表：%s（一张图看完关键页）" % sheet)
            else:
                last = [l for l in (r.stderr or r.stdout).strip().splitlines() if l.strip()]
                print("✗ PDF 未生成：%s" % (last[-1] if last else "未知原因"))
                print("   这台机器上 Word 的脚本接口不稳定（常见 -2700）。")
                print("   需要看图请手动打开 docx → 文件 → 导出为 PDF，然后再跑一次本命令，")
                print("   preview 会自动复用那份 PDF 生成联系表。")
                print("   格式验收请直接用 verify.py，它不依赖任何渲染器。")
                ok = False

    print()
    if ok:
        print("可以直接交付。打开后如果 Word 提示更新域，选「是」，编号和目录才会刷新。")
        return 0
    print("有问题，见上面的 ✗ 项。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
