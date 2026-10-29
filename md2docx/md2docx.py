#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx.py —— 一条命令：Markdown → 排版好的 Word 报告

用法：
    python3 md2docx.py 报告.md
    python3 md2docx.py 报告.md 输出.docx
    python3 md2docx.py 报告.md out.docx --profile official
    python3 md2docx.py --list-profiles
    python3 md2docx.py 报告.md out.docx -- --toc --number-sections   # -- 之后原样传给 pandoc

做的事：
    1. 按 profile 编译（或复用）pandoc 的 reference.docx
    2. pandoc 转换（关代码高亮、关智能引号）
    3. docxfix 后处理：题注/引用/目录/标题编号/表格线框/分节页码
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "profiles")
BUILD_DIR = os.path.join(HERE, "build")
DEFAULT_PROFILE = "technical"


def die(msg):
    print(msg, file=sys.stderr)
    sys.exit(1)


def load_profile_json(path):
    """读 profile（容忍 // 和 /* */ 注释）"""
    sys.path.insert(0, HERE)
    from docxfix import load_profile
    return load_profile(path)


def list_profiles():
    if not os.path.isdir(PROFILE_DIR):
        die("找不到 profiles 目录：%s" % PROFILE_DIR)
    print("可用 profile：")
    for f in sorted(os.listdir(PROFILE_DIR)):
        if not f.endswith(".json"):
            continue
        name = f[:-5]
        try:
            label = load_profile_json(os.path.join(PROFILE_DIR, f)).get("name", "")
        except Exception:
            label = ""
        print("  %-14s %s" % (name, label))


def resolve_profile(arg):
    """--profile 可以给名字、给路径、给相对路径"""
    if not arg:
        return DEFAULT_PROFILE, os.path.join(PROFILE_DIR, DEFAULT_PROFILE + ".json")
    for c in cands_of(arg):
        if os.path.isfile(c):
            return os.path.basename(c).replace(".json", ""), os.path.abspath(c)
    die("找不到 profile：%s（用 --list-profiles 看看）" % arg)


def cands_of(arg):
    return [arg,
            os.path.join(PROFILE_DIR, arg),
            os.path.join(PROFILE_DIR, arg + ".json"),
            os.path.join(HERE, arg),
            os.path.join(HERE, arg + ".json"),
            os.path.join(os.getcwd(), arg),
            os.path.join(os.getcwd(), arg + ".json")]


def ensure_template(name, profile_path, reference=None):
    """每次都重新编译模板到临时目录。

    以前会缓存在 build/<name>.docx，有两个毛病：
      1) 不同项目用同名 profile 会互相覆盖（串格式）
      2) 手改过缓存的 docx 会在下次编译时被静默冲掉
    现在改为每次现编（0.2 秒），行为可预期。
    """
    builder = os.path.join(HERE, "build_template.py")
    tmpdir = tempfile.mkdtemp(prefix="md2docx-tpl.")
    tpl = os.path.join(tmpdir, name + ".docx")
    cmd = [sys.executable, builder, "-p", profile_path, "-o", tpl]
    if reference:
        cmd += ["--base", reference]
    r = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True)
    if r.returncode != 0:
        shutil.rmtree(tmpdir, ignore_errors=True)
        die("模板编译失败：\n" + (r.stderr or r.stdout))
    # 编译日志里的「✓ 生成 /var/folders/...」是临时路径，会让人误以为是成品，
    # 只保留有信息量的行（新建样式、主题色转换等）。
    for line in (r.stdout or "").splitlines():
        if line.strip() and not line.strip().startswith("✓"):
            print(line.rstrip())
    if r.stderr.strip():
        print(r.stderr.strip(), file=sys.stderr)
    return tpl, tmpdir


def highlight_flag():
    """pandoc 3.x 用 --syntax-highlighting，2.x 用 --no-highlight"""
    try:
        out = subprocess.run(["pandoc", "--version"], capture_output=True, text=True).stdout
    except FileNotFoundError:
        die("没找到 pandoc，请先安装：https://pandoc.org/installing.html")
    m = re.search(r"pandoc (\d+)", out)
    major = int(m.group(1)) if m else 0
    return "--syntax-highlighting=none" if major >= 3 else "--no-highlight"


def read_frontmatter(md_path):
    """读 md 的 YAML 头（只取我们关心的几个键，不需要装 PyYAML）"""
    meta = {}
    try:
        with open(md_path, encoding="utf-8") as fh:
            head = fh.read(4000)
    except OSError:
        return meta
    if not head.startswith("---"):
        return meta
    end = head.find("\n---", 3)
    if end < 0:
        return meta
    block = head[3:end]
    for line in block.splitlines():
        m = re.match(r"^([A-Za-z_][A-Za-z0-9_-]*):[ \t]*(.*)$", line)
        if not m:
            continue
        k, v = m.group(1), m.group(2).strip()
        if not v or v == "|" or v == ">":
            continue
        if v.startswith("[") and v.endswith("]"):
            # 行内列表：keywords: [甲, 乙] / [甲，乙]
            inner = v[1:-1].strip()
            if not inner:
                meta[k] = []
            else:
                items = [x.strip().strip("'\"") for x in
                         re.split(r"[,，;；]", inner) if x.strip()]
                meta[k] = items
            continue
        v = v.strip(chr(39) + chr(34))
        if v:
            meta[k] = v
    return meta


def citation_args(md_path, profile):
    """参考文献：--citeproc + --csl

    - md 头部写 bibliography: refs.bib （或在 profile 里配）
    - CSL 风格：内置 csl/ 目录下那几套国标 GB/T 7714
    """
    cit = (profile or {}).get("citations") or {}
    meta = read_frontmatter(md_path)
    bib = meta.get("bibliography") or cit.get("bibliography")
    if not bib:
        return []
    bib_path = bib if os.path.isabs(bib) else os.path.join(os.path.dirname(md_path), bib)
    if not os.path.exists(bib_path):
        print("! 找不到参考文献文件 %s，跳过引用处理" % bib_path, file=sys.stderr)
        return []
    args = ["--citeproc", "--bibliography", bib_path]
    # 角标可点击跳转到后面的文献条目：pandoc 会给每条文献加 ref-xxx 书签，
    # 正文里的 [1] 变成指向它的内部链接。不开启则只是纯文字。
    if cit.get("link_citations", True):
        args.append("-M")
        args.append("link-citations=true")
    csl = meta.get("csl") or cit.get("csl") or "chinese-gb7714-2005-numeric"
    if csl.endswith(".csl"):
        csl_path = csl if os.path.isabs(csl) else os.path.join(
            os.path.dirname(md_path), csl)
    else:
        csl_path = os.path.join(HERE, "csl", csl + ".csl")
    if os.path.exists(csl_path):
        args += ["--csl", csl_path]
    else:
        print("! 找不到 CSL 样式 %s，用 pandoc 默认样式" % csl_path, file=sys.stderr)
    return args


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if "--list-profiles" in argv:
        list_profiles()
        return 0

    profile_arg = None
    reference = None
    if "--profile" in argv:
        i = argv.index("--profile")
        if i + 1 >= len(argv):
            die("--profile 后面要跟档案名")
        profile_arg = argv[i + 1]
        del argv[i:i + 2]
    elif any(a.startswith("--profile=") for a in argv):
        profile_arg = [a for a in argv if a.startswith("--profile=")][0].split("=", 1)[1]
        argv = [a for a in argv if not a.startswith("--profile=")]

    if "--reference" in argv:
        i = argv.index("--reference")
        if i + 1 >= len(argv):
            die("--reference 后面要跟模板 docx 的路径")
        reference = os.path.abspath(argv[i + 1])
        del argv[i:i + 2]
        if not os.path.isfile(reference):
            die("找不到模板文件：%s" % reference)
        print("· 使用自定义模板：%s" % reference)

    # `--` 之后是透传给 pandoc 的参数
    extra = []
    if "--" in argv:
        i = argv.index("--")
        extra = argv[i + 1:]
        argv = argv[:i]
    if not argv:
        die("用法：python3 md2docx.py <输入.md> [输出.docx] [--profile 名字] [-- pandoc 参数...]")

    md = argv[0]
    if not os.path.isfile(md):
        die("找不到输入文件：%s" % md)
    # 输出路径按调用者的当前目录解析（切换 pandoc cwd 后相对路径会跑偏）
    out = os.path.abspath(argv[1] if len(argv) >= 2
                          else os.path.splitext(md)[0] + ".docx")
    out_dir = os.path.dirname(out)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    # pandoc 在 md 所在目录里跑，这样 md 里的相对图片路径才找得到
    md_abs = os.path.abspath(md)
    md_dir = os.path.dirname(md_abs)
    md_name = os.path.basename(md_abs)
    extra = [a for a in argv[2:] if a.startswith("-")] + extra

    name, profile_path = resolve_profile(profile_arg)
    profile = load_profile_json(profile_path)
    tpl, tpldir = ensure_template(name, profile_path, reference)

    tmpdir = tempfile.mkdtemp(prefix="md2docx.")
    raw = os.path.join(tmpdir, "raw.docx")
    try:
        cmd = ["pandoc", md_name, "-o", raw,
               "--reference-doc", tpl,
               highlight_flag(),
               "-f", "markdown-smart"] + citation_args(md_abs, profile) + extra
        r = subprocess.run(cmd, cwd=md_dir, capture_output=True, text=True)
        if r.returncode != 0:
            die("pandoc 失败：\n" + (r.stderr or r.stdout))
        if r.stderr.strip():
            print(r.stderr.strip())

        # 把 YAML 里的 abstract / keywords 传给 docxfix（摘要需要单独排一页）
        meta = read_frontmatter(md_abs)
        meta_args = []
        kw = meta.get("keywords")
        if kw:
            meta_args = ["--meta", json.dumps({"keywords": kw}, ensure_ascii=False)]
        r = subprocess.run([sys.executable, os.path.join(HERE, "docxfix.py"),
                            raw, out, "--profile", profile_path] + meta_args,
                           cwd=HERE, text=True)
        if r.returncode != 0:
            die("docxfix 失败")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        shutil.rmtree(tpldir, ignore_errors=True)

    print()
    print("成品：%s" % out)
    try:
        sys.path.insert(0, HERE)
        import format_report
        rep = format_report.report(profile_path)
        keep = ("正文", "标题1", "图题注", "表格　", "页眉", "目录　",
                "摘要", "参考文献", "封面　")
        print("本次格式：")
        for line in rep.splitlines():
            if line.startswith(keep):
                print("  " + line.rstrip())
    except Exception:
        print("格式档案：%s" % profile_path)
    if (profile.get("header") or {}).get("even_text"):
        print("提示：已启用奇偶页页眉。Word 会让每一节都从奇数页开始，"
              "所以封面/目录/正文之间会多出空白页（双面印刷时这正是想要的）。")
        print("      不想留白页就把 header.even_text 清空。")
    print("在 Word 里按 Ctrl+A 再 F9（或右键目录→更新域），编号与目录会刷新。")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
