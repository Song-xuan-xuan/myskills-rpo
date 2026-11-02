#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx —— Markdown 报告转成排版规范、编号可自动重排的 Word

一个入口，六个子命令：

    python3 md2docx.py 报告.md [输出.docx] [--profile 档案]     # 转换（默认）
    python3 md2docx.py make 报告.md [--profile|--profiles ...]  # 转换 + 验收 + 摘要
    python3 md2docx.py batch 报告.md --profiles a,b,c           # 一份 md 出多版
    python3 md2docx.py verify 报告.docx [--json]                # 交付验收
    python3 md2docx.py preview 报告.docx [--sheet|--pages 1,3]  # 导 PDF / 联系表
    python3 md2docx.py styles 报告.docx                         # 逐段看样式（排查）
    python3 md2docx.py --list-profiles                          # 看有哪些档案

格式全在 profiles/*.json 里，用 configure.py 生成或修改。
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from _common import (HERE, PROFILE_DIR, DEFAULT_PROFILE, die, run,
                     strip_json_comments, resolve_profile, load_profile)

# ============================================================
# ① 转换：md -> docx（pandoc + build_template + docxfix）
# ============================================================

def list_profiles():
    if not os.path.isdir(PROFILE_DIR):
        die("找不到 profiles 目录：%s" % PROFILE_DIR)
    print("可用 profile：")
    for f in sorted(os.listdir(PROFILE_DIR)):
        if not f.endswith(".json"):
            continue
        name = f[:-5]
        try:
            label = load_profile(os.path.join(PROFILE_DIR, f)).get("name", "")
        except Exception:
            label = ""
        print("  %-14s %s" % (name, label))


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


def cmd_convert(argv):
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

    name, profile_path = _resolve_or_die(profile_arg)
    profile = load_profile(profile_path)
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
        import configure as _cfg
        rep = _cfg.report(profile_path)
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


# ============================================================
# ② 批量：一份 md 出多个版本
# ============================================================

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


def cmd_batch(argv):
    argv = list(argv)
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
        cmd = [sys.executable, __file__, "convert", md, out,
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


# ============================================================
# ③ 验收：能程序判的全判掉
# ============================================================

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def wq(t):
    return "{%s}%s" % (W[1:-1], t)


# ---------------------------------------------------------------- 工具

def load_docxfix():
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "dx", os.path.join(HERE, "docxfix.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def style_of(p):
    ps = p.find(wq("pPr") + "/" + wq("pStyle"))
    return ps.get(wq("val")) if ps is not None else None


def para_text(p):
    return "".join(t.text or "" for t in p.iter(wq("t"))).strip()


def has_child(p, path):
    cur = p
    for tag in path.split("/"):
        if cur is None:
            return None
        cur = cur.find(wq(tag))
    return cur


def style_has(path, tag):
    """哪些样式里带某个 pPr 子元素（含 basedOn 继承）"""
    z = zipfile.ZipFile(path)
    try:
        sty = z.read("word/styles.xml").decode("utf-8")
    except KeyError:
        return set()
    import xml.etree.ElementTree as _ET
    root = _ET.fromstring(sty.encode("utf-8"))
    direct, based = {}, {}
    for st in root.findall(wq("style")):
        sid = st.get(wq("styleId"))
        b = st.find(wq("basedOn"))
        based[sid] = b.get(wq("val")) if b is not None else None
        pPr = st.find(wq("pPr"))
        direct[sid] = (pPr is not None and pPr.find(wq(tag)) is not None)
    out = set()
    for sid in direct:
        cur, seen = sid, set()
        while cur and cur not in seen:
            seen.add(cur)
            if direct.get(cur):
                out.add(sid)
                break
            cur = based.get(cur)
    return out


def para_keeps_next(p, keep_styles):
    if has_child(p, "pPr/keepNext") is not None:
        return True
    return style_of(p) in keep_styles


def instr_texts(p):
    return [i.text or "" for i in p.iter(wq("instrText"))]


# ---------------------------------------------------------------- 各项检查

def is_template(path):
    """模板/样张不是交付物：没有域、段落很少"""
    try:
        d = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    except Exception:
        return False
    if "<w:instrText" in d:
        return False
    return d.count("<w:p ") < 40


def check_package(path, problems, warnings):
    """结构合规：交给 docxfix.validate（元素顺序、跨部件引用、书签配对等）"""
    dx = load_docxfix()
    from docx import Document
    doc = Document(path)
    for msg in dx.validate(doc, quiet=True):
        problems.append(msg)
    return doc


def check_captions(doc, path, problems, warnings, stats):
    """题注：必须有 SEQ 域、有文字；图/表题注要跟主体同页"""
    body = doc.element.body
    kids = list(body)
    keep_styles = style_has(path, "keepNext")
    fig_n = tbl_n = eq_n = 0
    seq_fig = seq_tbl = seq_eq = 0
    empty = []
    not_same_page = []

    # 封面上的 logo 也是 drawing，不能算成「图」。
    # 正文从第一个标题段落开始，之前的都当封面/目录。
    body_start = 0
    for i, node in enumerate(kids):
        if node.tag == wq("p") and (style_of(node) or "").startswith("Heading"):
            body_start = i
            break

    for i, node in enumerate(kids):
        if node.tag != wq("p") and node.tag != wq("tbl"):
            continue
        sid = style_of(node) if node.tag == wq("p") else None

        if (node.tag == wq("p") and i >= body_start
                and node.find(".//" + wq("drawing")) is not None):
            fig_n += 1
            # 图段落应该有 keepNext，且下一个段落是图注
            keep = para_keeps_next(node, keep_styles)
            nxt = kids[i + 1] if i + 1 < len(kids) else None
            is_cap = (nxt is not None and nxt.tag == wq("p")
                      and style_of(nxt) in ("ImageCaption", "Caption"))
            if not keep or not is_cap:
                not_same_page.append(para_text(nxt) if nxt is not None else "(图后无内容)")
        if node.tag == wq("tbl"):
            tbl_n += 1
            nxt = kids[i + 1] if i + 1 < len(kids) else None
            prev = kids[i - 1] if i > 0 else None
            near = [x for x in (prev, nxt) if x is not None and x.tag == wq("p")
                    and style_of(x) in ("TableCaption", "Caption")]
            if near:
                # 表注在表前才要 keepNext；在表后无所谓
                if prev is not None and style_of(prev) in ("TableCaption", "Caption"):
                    if not para_keeps_next(prev, keep_styles):
                        not_same_page.append(para_text(prev))
        if node.tag == wq("p") and node.find(M + "oMathPara") is not None:
            eq_n += 1

        if sid in ("ImageCaption", "TableCaption"):
            text = para_text(node)
            had_seq = any("SEQ" in x for x in instr_texts(node))
            if any("SEQ 图" in x for x in instr_texts(node)):
                seq_fig += 1
            if any("SEQ 表" in x for x in instr_texts(node)):
                seq_tbl += 1
            # 去掉编号后还剩多少字
            rest = re.sub(r"^(图|表)\s*\d*[\s　]*", "", text)
            if not rest:
                empty.append(text or "(空)")
            if not had_seq:
                warnings.append("题注没用 SEQ 域，增删图表不会自动编号：%s" % text[:20])

    stats.update({"图": fig_n, "表": tbl_n, "公式": eq_n})
    if fig_n and seq_fig < fig_n:
        problems.append("图 %d 个，但有编号的图注只有 %d 个" % (fig_n, seq_fig))
    if tbl_n and seq_tbl < tbl_n:
        problems.append("表 %d 个，但有编号的表注只有 %d 个" % (tbl_n, seq_tbl))
    if empty:
        warnings.append("%d 个题注只有编号没有文字：%s" % (len(empty), "、".join(empty[:3])))
    if not_same_page:
        warnings.append("图/表与其题注可能被分页拆开（缺 keepNext）：%s"
                        % "、".join(x[:16] for x in not_same_page[:3]))


def check_refs(doc, problems, warnings):
    """交叉引用：每个 REF 都要有落点"""
    body = doc.element.body
    marks = {b.get(wq("name")) for b in body.iter(wq("bookmarkStart"))
             if b.get(wq("name"))}
    targets = set()
    for node in body.iter(wq("p")):
        for code in instr_texts(node):
            m = re.search(r"(?<![A-Za-z])REF\s+([\w\-.:]+)", code)
            if m:
                targets.add(m.group(1))
    missing = sorted(t for t in targets if t not in marks)
    if missing:
        problems.append("交叉引用找不到落点（F9 后会变「错误！未找到引用源」）：%s"
                        % "、".join(missing[:5]))
    dup = [n for n in marks if list(marks).count(n) > 1]
    if dup:
        warnings.append("锚点重名：%s" % "、".join(sorted(set(dup))[:3]))
    return len(targets)


def check_toc(doc, path, problems, warnings, cfg):
    """目录：域在不在、条目样式有没有制表位"""
    body = doc.element.body
    has_toc = any("TOC" in code for code in
                  (c for p in body.iter(wq("p")) for c in instr_texts(p)))
    z = zipfile.ZipFile(path)
    sty = z.read("word/styles.xml").decode("utf-8")
    toc_blocks = re.findall(r'<w:style [^>]*w:styleId="(TOC\d)".*?</w:style>',
                            sty, re.S)
    if not has_toc:
        # 只有文档里真写了「目录」这个标题，才说明用户要目录却没生成
        want = any(para_text(p).strip() in ("目　录", "目录", "Contents", "CONTENTS")
                   for p in body.iter(wq("p")))
        if want:
            warnings.append("有「目录」标题但没有 TOC 域，F9 也生成不出条目")
    else:
        used = [m for m in re.finditer(
            r'<w:style [^>]*w:styleId="(TOC\d)".*?</w:style>', sty, re.S)]
        no_tab = [m.group(1) for m in used if "<w:tabs>" not in m.group(0)]
        if no_tab:
            problems.append("目录条目样式缺制表位，页码不会右对齐：%s"
                            % ",".join(no_tab[:4]))
        no_ind = [m.group(1) for m in used
                  if m.group(1) != "TOC1" and 'w:leftChars="0"' in m.group(0)]
        if no_ind:
            warnings.append("目录条目没有层级缩进：%s" % ",".join(no_ind[:3]))
    return has_toc


def check_bibliography(doc, problems, warnings):
    """参考文献：编号必须是 SEQ 域，正文角标必须是 REF 域"""
    body = doc.element.body
    entries, pending = [], []
    for node in list(body):
        if node.tag == wq("bookmarkStart"):
            nm = node.get(wq("name")) or ""
            if nm.startswith("ref-"):
                pending.append(nm)
            continue
        if node.tag == wq("p") and style_of(node) == "Bibliography":
            entries.append(node)
        pending = []
    if not entries:
        return 0
    static = 0
    for p in entries:
        text = para_text(p)
        codes = " ".join(instr_texts(p))
        if re.match(r"^\[\s*\d+\s*\]", text) and "SEQ" not in codes:
            static += 1
    if static:
        warnings.append("参考文献编号是静态文字（%d 条），删条目后 F9 不会重排" % static)
    # 角标
    anchors = [h.get(wq("anchor")) or "" for h in body.iter(wq("hyperlink"))]
    cite_links = [a for a in anchors if a.startswith("ref-")]
    static_cite = 0
    for p in body.iter(wq("p")):
        codes = " ".join(instr_texts(p))
        if "REF " in codes or "seq" in codes.lower():
            continue
        for h in p.iter(wq("hyperlink")):
            if (h.get(wq("anchor")) or "").startswith("ref-"):
                if "REF" not in " ".join(instr_texts(h)):
                    static_cite += 1
    if static_cite:
        warnings.append("%d 处正文角标是静态数字，删条目后不会跟着变" % static_cite)
    return len(entries)


def check_footer(doc, path, problems, warnings):
    """页脚：要有 PAGE 域"""
    z = zipfile.ZipFile(path)
    footers = [n for n in z.namelist() if re.match(r"word/footer\d*\.xml$", n)]
    if not footers:
        warnings.append("没有页脚（配置里 footer.page_number=false 时可忽略）")
        return
    codes = " ".join("".join(re.findall(r"<w:instrText[^>]*>([^<]*)<",
                                        z.read(n).decode("utf-8")))
                     for n in footers)
    if "PAGE" not in codes:
        problems.append("页脚里没有页码域")


def check_images(doc, path, problems, warnings):
    """图片：不能超版心；有效分辨率别太低"""
    try:
        sec = doc.sections[0]
        limit = int(sec.page_width - sec.left_margin - sec.right_margin)
    except Exception:
        return
    body = doc.element.body
    over = 0
    low_dpi = []
    z = zipfile.ZipFile(path)
    media = {n: z.read(n) for n in z.namelist() if n.startswith("word/media/")}
    try:
        from PIL import Image
        import io
    except ImportError:
        Image = None

    for ext in body.iter(WP + "extent"):
        cx = int(ext.get("cx") or 0)
        if cx > limit + 5000:
            over += 1
    if over:
        problems.append("%d 张图片超出版心宽度" % over)

    if Image is not None:
        # 图片按显示尺寸算有效 dpi
        sizes = {}
        for name, blob in media.items():
            try:
                im = Image.open(io.BytesIO(blob))
                sizes[name] = im.size
            except Exception:
                pass
        dn = 0
        for ext in body.iter(WP + "extent"):
            cx = int(ext.get("cx") or 0)
            if cx <= 0:
                continue
            names = sorted(sizes)
            if dn < len(names):
                px_w = sizes[names[dn]][0]
                inches = cx / 914400.0
                dpi = px_w / inches if inches else 0
                if 0 < dpi < 150:
                    low_dpi.append(int(dpi))
            dn += 1
        if low_dpi:
            warnings.append("图片有效分辨率偏低（%s dpi），打印可能发虚"
                            % "、".join(str(x) for x in sorted(low_dpi)[:3]))


def check_tables(doc, cfg, problems, warnings):
    """表格：列数一致、长表表头重复、行不被拆"""
    body = doc.element.body
    n = 0
    for tbl in body.iter(wq("tbl")):
        n += 1
        grid = tbl.find(wq("tblGrid"))
        ncol = len(grid.findall(wq("gridCol"))) if grid is not None else 0
        rows = tbl.findall(wq("tr"))
        bad = 0
        for tr in rows:
            cells = tr.findall(wq("tc"))
            span = 0
            for tc in cells:
                gs = tc.find(wq("tcPr") + "/" + wq("gridSpan"))
                span += int(gs.get(wq("val")) or 1) if gs is not None else 1
            if ncol and span != ncol:
                bad += 1
        if bad:
            warnings.append("有表格的列数与网格不一致（%d 行），换行可能错位" % bad)
        head = rows[0].find(wq("trPr") + "/" + wq("tblHeader")) if rows else None
        if len(rows) > 12 and head is None:
            warnings.append("有 %d 行的表格没设「跨页重复表头」" % len(rows))
        if len(rows) > 25 and not all(
                tr.find(wq("trPr") + "/" + wq("cantSplit")) is not None
                for tr in rows):
            warnings.append("有 %d 行的超长表格，行可能被拆到两页" % len(rows))
    return n


def check_numbering(doc, path, problems, warnings):
    """标题自动编号：样式或段落上要有 numPr"""
    z = zipfile.ZipFile(path)
    sty = z.read("word/styles.xml").decode("utf-8")
    heads = [m for m in re.finditer(
        r'<w:style [^>]*w:styleId="Heading([1-6])".*?</w:style>', sty, re.S)]
    if not heads:
        return
    unnumbered = [m.group(1) for m in heads
                  if "<w:numPr>" not in m.group(0) and "w:numId" not in m.group(0)]
    paras = [p for p in doc.element.body.iter(wq("p"))
             if (style_of(p) or "").startswith("Heading")]
    para_num = sum(1 for p in paras
                   if has_child(p, "pPr/numPr") is not None)
    if para_num < len(paras) and len(unnumbered) >= len(heads):
        warnings.append("标题没有自动编号（配置里 heading_numbering.enabled=false 时可忽略）")
    elif para_num and para_num < len(paras):
        warnings.append("%d 个标题段落没有编号（可能是附录之外的手写标题）"
                        % (len(paras) - para_num))


def check_fields(doc, problems, warnings):
    """域的整体健康度"""
    body = doc.element.body
    n_seq = n_ref = 0
    bad_field = 0
    for p in body.iter(wq("p")):
        codes = instr_texts(p)
        for c in codes:
            if "SEQ" in c:
                n_seq += 1
            if re.search(r"(?<![A-Za-z])REF\s", c):
                n_ref += 1
        # 域必须成对：数 begin / end
        begins = len(list(p.iter(wq("fldChar"))))
    # 全局统计 fldChar begin/end 配对
    begins = len(re.findall(r'<w:fldChar w:fldCharType="begin"',
                            ET.tostring(body, encoding="unicode")))
    ends = len(re.findall(r'<w:fldChar w:fldCharType="end"',
                          ET.tostring(body, encoding="unicode")))
    if begins != ends:
        problems.append("域的开始/结束数量不配对（begin %d / end %d），Word 可能报错"
                        % (begins, ends))
    return {"SEQ": n_seq, "REF": n_ref}


def check_text_hygiene(doc, problems, warnings):
    """纯文本层面一眼能看出来的毛病"""
    body = doc.element.body
    full = "".join(t.text or "" for t in body.iter(wq("t")))
    if "{{" in full or "}}" in full:
        bad = re.findall(r"\{\{[^}]{0,30}\}\}", full)
        problems.append("文档里还有没替换的交叉引用标记：%s" % "、".join(bad[:3]))
    if re.search(r"关键词[：:]\s*\[", full):
        warnings.append("关键词带方括号")
    if "？" in full and "?" not in full:
        pass
    # 空白过多（可能出现空段）
    empties = sum(1 for p in body.iter(wq("p"))
                  if not para_text(p) and p.find(wq("pPr")) is None)
    if empties > 30:
        warnings.append("文档里有 %d 个完全空的段落" % empties)


# ---------------------------------------------------------------- 主流程

def verify(path, as_json=False):
    problems, warnings = [], []
    stats, fields = {}, {}
    doc = check_package(path, problems, warnings)

    cfg = {}    # 不再读侧车文件；需要的判断都从文档本身推断

    check_captions(doc, path, problems, warnings, stats)
    check_refs(doc, problems, warnings)
    check_toc(doc, path, problems, warnings, cfg)
    n_bib = check_bibliography(doc, problems, warnings)
    if n_bib:
        stats["文献"] = n_bib
    check_footer(doc, path, problems, warnings)
    check_images(doc, path, problems, warnings)
    check_tables(doc, cfg, problems, warnings)
    check_numbering(doc, path, problems, warnings)
    fields = check_fields(doc, problems, warnings)
    check_text_hygiene(doc, problems, warnings)

    # 真正需要眼睛的项
    eyes = []
    if stats.get("图") or stats.get("表"):
        eyes.append("图/表与题注是否同页（需渲染后由人判断）")
    if fields.get("REF"):
        eyes.append("交叉引用编号是否指向正确位置（需渲染后由人判断）")
    if fields.get("SEQ"):
        eyes.append("增删图表后编号是否连续（需渲染后由人判断）")

    result = {
        "path": path,
        "stats": stats,
        "fields": fields,
        "problems": problems,
        "warnings": warnings,
        "need_eyes": eyes,
        "ok": not problems,
    }
    if as_json:
        return result

    print("=" * 64)
    print(os.path.basename(path))
    print("=" * 64)
    if stats:
        print("内容：%s" % " · ".join("%s %d" % (k, v) for k, v in stats.items()
                                   if v))
    if fields:
        print("域：  %s" % " · ".join("%s %d" % (k, v) for k, v in fields.items()
                                   if v))
    if problems:
        print("\n[错误] %d 项" % len(problems))
        for x in problems:
            print("  ✗ " + x)
    if warnings:
        print("\n[提醒] %d 项" % len(warnings))
        for x in warnings:
            print("  · " + x)
    print()
    if not problems:
        print("程序能判的都通过了。")
    if eyes:
        print("以下几项程序判不了，需要人看（不要自己转 PDF 逐页截图）：")
        for x in eyes:
            print("  - " + x)
        print("  做法：把 docx 路径交给用户，让他在 Word 里 Ctrl+A → F9 后自己确认。")
        print("  确实要自己看，只能用：preview.py 报告.docx --sheet（一张联系表）")
    else:
        print("没有需要肉眼确认的项。")
    return result


def expand(args):
    out = []
    for a in args:
        if os.path.isdir(a):
            out.extend(glob.glob(os.path.join(a, "**", "*.docx"), recursive=True))
        elif any(ch in a for ch in "*?"):
            out.extend(glob.glob(a, recursive=True))
        else:
            out.append(a)
    seen, uniq = set(), []
    for p in out:
        ap = os.path.abspath(p)
        if ap in seen:
            continue
        if not os.path.isfile(ap) or os.path.basename(ap).startswith("~$"):
            continue
        seen.add(ap)
        uniq.append(ap)
    return uniq


def cmd_verify(argv):
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]
    if not args:
        print(__doc__)
        return 1
    files = expand(args)
    if not files:
        print("没有找到 docx")
        return 1

    as_json = "--json" in flags
    results = [verify(f, as_json=True) for f in files]
    if as_json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for f in files:
            verify(f)

    bad = [r for r in results if not r["ok"]]
    if not as_json:
        print()
        print("=" * 64)
        print("%d 个文件，%d 个有错误" % (len(results), len(bad)))
    return 1 if bad else 0


# ============================================================
# ④ 预览：导 PDF，或把关键页拼成一张联系表
# ============================================================

HERE = os.path.dirname(os.path.abspath(__file__))
WORD_CONTAINER = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Word/Data/Documents")
WORD_WORK = os.path.join(WORD_CONTAINER, ".md2docx_preview")

# Word 只有一个实例：两个会话同时驱动会互相干扰（open/active document 打架）。
# 用文件锁排队，避免「找不到活动文档」「转出来是别的文件」这类怪现象。
LOCK_PATH = os.path.join(tempfile.gettempdir(), "md2docx-word.lock")


def acquire_lock(timeout=240):
    try:
        import fcntl
    except ImportError:
        return None
    fh = open(LOCK_PATH, "w")
    waited = 0
    while True:
        try:
            fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
            return fh
        except BlockingIOError:
            if waited == 0:
                print("另一个会话正在用 Word，排队等待…", file=sys.stderr)
            if waited >= timeout:
                fh.close()
                raise SystemExit("等了 %d 秒还没轮上，请稍后再试" % timeout)
            time.sleep(2)
            waited += 2


def release_lock(fh):
    if fh is None:
        return
    try:
        import fcntl
        fcntl.flock(fh, fcntl.LOCK_UN)
    except Exception:
        pass
    try:
        fh.close()
    except Exception:
        pass

CONVERT = '''
on run argv
    set src to item 1 of argv
    set dst to item 2 of argv
    set wantName to item 3 of argv
    with timeout of 900 seconds
        tell application "Microsoft Word"
            activate
            -- Word 只有一个实例，别人也在用时 open 可能被吞掉。
            -- 所以：反复尝试打开，并在文档列表里按文件名确认真的打开了。
            set myDoc to missing value
            repeat 6 times
                try
                    open file name src
                end try
                delay 2
                repeat with d in (get documents)
                    try
                        if (name of d) is wantName then
                            set myDoc to contents of d
                            exit repeat
                        end if
                    end try
                end repeat
                if myDoc is not missing value then exit repeat
            end repeat
            if myDoc is missing value then
                error "打不开 " & wantName & "（Word 正被其它任务占用？）"
            end if
            try
                update fields of text object of myDoc
            end try
            delay 1
            save as myDoc file name dst file format format PDF
        end tell
    end timeout
    return "ok"
end run
'''

COUNT_DOCS = '''
tell application "Microsoft Word"
    try
        return (count of documents) as string
    on error
        return "0"
    end try
end tell
'''

QUIT_WORD = '''
tell application "Microsoft Word"
    try
        quit saving no
    end try
end tell
'''


def run_osa(script, *args):
    with tempfile.NamedTemporaryFile("w", suffix=".applescript",
                                     delete=False, encoding="utf-8") as fh:
        fh.write(script)
        path = fh.name
    try:
        return subprocess.run(["osascript", path] + list(args),
                              capture_output=True, text=True)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def dismiss_dialogs():
    """关掉 Word 上的模态弹窗

    Word 有弹窗时（例如「文件可能已经损坏」「此文件被锁定」），
    所有 AppleScript 的 open 都会被静默忽略，表现为「打不开但也不报错」。
    """
    script = '''
tell application "System Events"
    tell process "Microsoft Word"
        set n to 0
        repeat 5 times
            set did to false
            try
                repeat with w in windows
                    try
                        if (count of sheets of w) > 0 then
                            set b to button 1 of sheet 1 of w
                            click b
                            set did to true
                            delay 0.5
                        end if
                    end try
                end repeat
            end try
            if did then set n to n + 1
            if not did then exit repeat
        end repeat
        return n as string
    end tell
end tell
'''
    r = run_osa(script)
    return (r.stdout or "").strip()


def word_state():
    """返回 Word 的进程状态：是否运行、PID、PPID、窗口数、文档数

    为什么要看 PPID：如果 Word 是被脚本拉起来的，脚本退出后 Word 的父进程
    变成 1（孤儿化），这种实例常常「进程活着、能答属性查询，但没有窗口，
    open 被静默丢弃」——看起来像僵尸，实际上是 GUI 会话没起来。
    """
    st = {"running": False, "pid": None, "ppid": None,
          "windows": None, "docs": None}
    r = subprocess.run(["pgrep", "-x", "Microsoft Word"], capture_output=True, text=True)
    pid = (r.stdout or "").strip().splitlines()
    if not pid:
        return st
    st["running"] = True
    st["pid"] = pid[0].strip()
    try:
        ps = subprocess.run(["ps", "-o", "ppid=", "-p", st["pid"]],
                            capture_output=True, text=True).stdout.strip()
        st["ppid"] = int(ps) if ps.isdigit() else None
    except Exception:
        pass
    try:
        st["docs"] = int(run_osa(COUNT_DOCS).stdout.strip() or 0)
    except ValueError:
        pass
    win = run_osa('''
tell application "System Events"
    tell process "Microsoft Word"
        try
            return (count of windows) as string
        on error
            return "0"
        end try
    end tell
end tell
''')
    try:
        st["windows"] = int((win.stdout or "0").strip() or 0)
    except ValueError:
        pass
    return st


def explain_word_state(st):
    """把 Word 状态翻译成「现在怎么办」"""
    if not st["running"]:
        print("Word：未运行。preview.py 不会替你冷启动它（冷启动容易卡在引导页）。")
        return "not-running"
    pid, ppid = st["pid"], st["ppid"]
    win, docs = st["windows"], st["docs"]
    print("Word：运行中（PID %s，父进程 %s），窗口 %s 个，文档 %s 个"
          % (pid, ppid, win, docs))
    if win == 0 and not docs:
        print()
        print("  诊断：Word 进程在，但没有窗口 —— 它被「孤儿化」了。")
        if ppid == 1:
            print("        父进程是 1，说明当初是脚本/后台拉起来的，脚本退出后它就没再建界面。")
        print("        这种实例能回答属性查询，但 open 会被静默丢弃（表现为 -2700）。")
        print()
        print("  处理（任选）：")
        print("    终端：kill -9 %s" % pid)
        print("    或：活动监视器 → 搜索 Microsoft Word → 选中 → 点 × → 强制退出")
        print("    然后从 Dock / 应用程序文件夹手动启动 Word，等它出现正常窗口。")
        return "orphaned"
    if win == 0:
        print("  诊断：Word 有文档没有窗口，可能卡在弹窗上。关掉弹窗再试。")
        return "no-window"
    return "ok"

def diagnose():
    """转换失败时给出可执行的诊断，而不是只丢一个错误码"""
    st = word_state()
    print("诊断：", file=sys.stderr)
    kind = explain_word_state(st)
    if kind == "ok":
        print("  → Word 状态正常但打不开文件，可能是另一个任务正在用它。",
              file=sys.stderr)
        print("    稍等重试，或直接双击 docx 自己看效果。", file=sys.stderr)
    elif kind == "not-running":
        print("  → 先手动启动一次 Word，确认没有弹窗，再重试 --word。",
              file=sys.stderr)


def reset_word():
    """Word 有时会进入异常状态（文档列表里出现 missing value），杀掉重启即可

    典型诱因：上次预览的临时文件被外部删除 / Word 卡在弹框上。
    """
    # 不能用 SIGKILL：macOS 会弹「执行了强制退出」错误报告，用户以为文件坏了。
    # 温和地让 Word 退出即可；退不掉就让用户自己处理。
    subprocess.run(["osascript", "-e",
                    'tell application "Microsoft Word" to quit saving no'],
                   capture_output=True, text=True)
    time.sleep(2)


# ---------------------------------------------------------------- 关键页 / 联系表

def pick_key_pages(pdf, total):
    """挑出最值得看的那几页：封面、目录、正文首页、最后一页"""
    pages = [1]
    try:
        out = subprocess.run(["pdftotext", "-layout", pdf, "-"],
                             capture_output=True, text=True).stdout
        texts = out.split("\f")
    except Exception:
        texts = []
    # 目录页：含「目 录」或大量点线
    for i, t in enumerate(texts[:8], start=1):
        if i == 1:
            continue
        if "目\u3000录" in t or "目录" in t or t.count("..") > 3:
            pages.append(i)
            break
    # 正文第一页：目录之后第一个有正文章节号的行
    start = pages[-1] + 1 if len(pages) > 1 else 2
    for i in range(start, min(total, start + 6) + 1):
        pages.append(i)
        break
    if total not in pages:
        pages.append(total)
    return sorted(set(p for p in pages if 1 <= p <= total))


def render_pages(pdf, pages, dpi=80):
    """把指定页渲染成 PNG，返回 [(页码, 路径)]"""
    out = []
    tmpdir = tempfile.mkdtemp(prefix="md2docx-pages.")
    for p in pages:
        base = os.path.join(tmpdir, "p%03d" % p)
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p),
                        "-l", str(p), pdf, base], capture_output=True)
        for f in os.listdir(tmpdir):
            if f.startswith("p%03d" % p) and f.endswith(".png"):
                out.append((p, os.path.join(tmpdir, f)))
    return out


def make_sheet(images, out_path, cols=2, label=True):
    """把若干页拼成一张联系表（一张图看完，省得逐页截图）"""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("没有 Pillow，改成分页导出（pip install pillow 可拼图）",
              file=sys.stderr)
        return None
    if not images:
        return None
    thumbs = []
    for num, f in images:
        try:
            im = Image.open(f).convert("RGB")
        except Exception:
            continue
        w = 900
        h = int(im.height * w / im.width)
        thumbs.append((num, im.resize((w, h), Image.LANCZOS)))
    if not thumbs:
        return None
    cols = max(1, min(int(cols), len(thumbs)))
    rows = (len(thumbs) + cols - 1) // cols
    pad, bar = 24, 34
    tw = max(t.width for _, t in thumbs)
    th = max(t.height for _, t in thumbs)
    W = cols * tw + (cols + 1) * pad
    H = rows * (th + bar) + (rows + 1) * pad
    sheet = Image.new("RGB", (W, H), (238, 238, 238))
    draw = ImageDraw.Draw(sheet)
    for i, (num, im) in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (th + bar + pad)
        sheet.paste(im, (x, y))
        draw.rectangle([x - 2, y - 2, x + im.width + 1, y + im.height + 1],
                       outline=(150, 150, 150))
        if label:
            # 用 ASCII 页码：PIL 内置字体没有中文字形，写「第 N 页」会变方框
            draw.text((x + 6, y + im.height + 8), "page %d" % num,
                      fill=(40, 40, 40))
    sheet.save(out_path)
    return out_path


def purge_leftovers(max_age_sec=0):
    """删掉本工具在 Word 容器里留下的临时文件

    这些东西留着会被 Word 认为「别人还在用」，下次打开就弹
    「此文件被锁定，锁者：另一用户」。max_age_sec=0 表示全部清掉。
    """
    n = 0
    if not os.path.isdir(WORD_WORK):
        return n
    now = time.time()
    for f in list(os.listdir(WORD_WORK)):
        if not (f.startswith("md2docx_") or f.startswith("~$md2docx_")):
            continue
        p = os.path.join(WORD_WORK, f)
        try:
            if max_age_sec and now - os.path.getmtime(p) < max_age_sec:
                continue
            os.remove(p)
            n += 1
        except OSError:
            pass
    return n


def close_our_document(basename):
    """转换完把刚打开的那个文档关掉

    Word 这版的 AppleScript 不支持 close（会报 -1708），但可以用
    Cmd+W 关掉最前面的窗口；关闭前先确认最前面那个确实是我们的文件。
    """
    script = '''
on run argv
    set wantName to item 1 of argv
    tell application "Microsoft Word"
        try
            if (name of active document) is not wantName then return "skip"
        on error
            return "skip"
        end try
        activate
    end tell
    delay 0.3
    tell application "System Events"
        tell process "Microsoft Word"
            keystroke "w" using command down
        end tell
    end tell
    return "closed"
end run
'''
    r = run_osa(script, basename)
    time.sleep(0.6)
    return (r.stdout or "").strip() == "closed"


def with_word(docx, _retry=True):
    """用 Word 渲染。文件放进 Word 自己的容器，避免弹授权框"""
    if not os.path.isdir(os.path.expanduser(
            "~/Library/Containers/com.microsoft.Word")):
        return None, None
    # 绝不「冷启动」Word：新启动的 Word 会弹引导页/登录页，
    # AppleScript 在那种状态下拿不到文档（-2700、document = missing value），
    # 而且会把 Word 留在一个卡死的状态里。
    if subprocess.run(["pgrep", "-x", "Microsoft Word"],
                      capture_output=True).returncode != 0:
        print("Word 当前没有运行。脚本不会替你启动它——冷启动的 Word 会卡在引导页。",
              file=sys.stderr)
        print("请先手动打开一次 Microsoft Word，确认没弹窗后关闭窗口（不必退出），再重试。",
              file=sys.stderr)
        return None, 0

    os.makedirs(WORD_WORK, exist_ok=True)

    # 转换前 Word 里开着几个文档（用于决定事后能不能退出 Word）
    try:
        n0 = int(run_osa(COUNT_DOCS).stdout.strip() or 0)
    except ValueError:
        n0 = 0

    # 每次用唯一文件名：Word 是单实例，多会话共用同一个名字会互相覆盖/读错文件
    # 名字必须全局唯一：Word 是单实例，两个会话用同名文件时，
    # 后打开的那个会看到「此文件被锁定，锁者：另一用户」。
    h = "%s_%d_%s" % (time.strftime("%H%M%S"), os.getpid(), uuid.uuid4().hex[:6])
    src = os.path.join(WORD_WORK, "md2docx_%s.docx" % h)
    dst = os.path.join(WORD_WORK, "md2docx_%s.pdf" % h)
    purge_leftovers()
    shutil.copy(docx, src)
    # 清掉可能残留的 Word owner 文件（~$xxx.docx）。
    # 不清的话 Word 会认为「此文件被锁定，锁者：另一用户」，直接拒绝打开。
    for junk in (os.path.join(WORD_WORK, "~$" + os.path.basename(src)),):
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except OSError:
                pass
    if os.path.exists(dst):
        os.remove(dst)

    dismiss_dialogs()
    r = run_osa(CONVERT, src, dst, os.path.basename(src))
    if r.returncode != 0 or not os.path.exists(dst):
        msg = (r.stderr or r.stdout).strip()
        for junk in (src, os.path.join(WORD_WORK, "~$" + os.path.basename(src))):
            try:
                os.remove(junk)
            except OSError:
                pass
        if _retry:
            print("Word 状态异常，重启后重试…", file=sys.stderr)
            reset_word()
            return with_word(docx, _retry=False)
        print("Word 转换失败：%s" % msg, file=sys.stderr)
        diagnose()
        return None, n0

    # 不把文档留在 Word 里：留一个就会让下一次打开同名/同路径文件弹「被锁定」
    close_our_document(os.path.basename(src))
    return dst, n0


def cmd_preview(argv):
    args = list(argv)
    flags = [a for a in args if a.startswith("--")]
    pos = [a for a in args if not a.startswith("--")]

    if "--doctor" in flags:
        print("Word 预览环境检查")
        print("-" * 46)
        st = word_state()
        explain_word_state(st)
        left = []
        if os.path.isdir(WORD_WORK):
            left = [f for f in os.listdir(WORD_WORK)
                    if f.startswith(("md2docx_", "~$md2docx_"))]
        print("本工具的临时文件：%d 个%s"
              % (len(left), "（可清理：preview.py --cleanup）" if left else ""))
        print("-" * 46)
        print("提醒：格式验收优先用 verify.py，它不需要渲染器。")
        return 0

    if "--cleanup" in flags:
        n = purge_leftovers()
        print("已清理 %d 个本工具留在 Word 容器里的临时文件" % n)
        print("如果 Word 里还有 md2docx_ 开头的文档，请手动关掉。")
        return 0

    if not pos or "-h" in flags or "--help" in flags:
        print(__doc__)
        return 1
    docx = os.path.abspath(pos[0])
    if not os.path.isfile(docx):
        die("找不到文件：%s" % docx)

    out = None
    if "--out" in args:          # 注意要在完整参数里找
        i = args.index("--out")
        if i + 1 < len(args):
            out = args[i + 1]

    # 已经有一份比 docx 新的 PDF 就直接用，不必再驱动 Word
    sibling = os.path.splitext(docx)[0] + ".pdf"
    reuse = None
    if (reuse is None and os.path.exists(sibling)
            and os.path.getmtime(sibling) >= os.path.getmtime(docx)
            and "--force" not in flags):
        reuse = sibling

    if reuse:
        pdf, n0 = reuse, 0
        print("复用已有 PDF：%s（不启动 Word）" % reuse)
    elif "--word" in flags:
        lock = acquire_lock()
        try:
            pdf, n0 = with_word(docx)
        finally:
            release_lock(lock)
    else:
        # 默认不自动驱动 Word：
        # Word 在 macOS 上只有一个实例，脚本驱动它很容易弹
        # 「此文件被锁定，锁者：另一用户」，而且这版 Word 的
        # AppleScript 不能可靠关闭文档，残留会越积越多。
        print("没有可用的 PDF，且默认不自动启动 Word。", file=sys.stderr)
        print("", file=sys.stderr)
        print("想要 PDF / 联系表，任选一种：", file=sys.stderr)
        print("  ① 自己双击打开 %s → 文件 → 导出为 PDF，"
              "再运行本命令（会自动复用那份 PDF）" % os.path.basename(docx),
              file=sys.stderr)
        print("  ② 明确允许驱动 Word：python3 preview.py \"%s\" --word"
              % os.path.basename(docx), file=sys.stderr)
        print("", file=sys.stderr)
        print("如果只是想确认格式对不对，先跑 verify.py，它不需要任何渲染器：",
              file=sys.stderr)
        print("  python3 verify.py \"%s\"" % os.path.basename(docx), file=sys.stderr)
        return 3
    if not pdf or not os.path.exists(pdf):
        die("转换失败")

    if out:
        final = os.path.abspath(out)
    else:
        final = os.path.join(os.path.dirname(docx),
                             os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if os.path.abspath(pdf) != os.path.abspath(final):
        shutil.copy(pdf, final)
    # 只删「Word 容器里的临时 PDF」。
    # 绝不能用 final（那可能是用户自己导出的 PDF，删了就丢数据）。
    if os.path.abspath(pdf).startswith(os.path.abspath(WORD_WORK) + os.sep):
        try:
            os.remove(pdf)
        except OSError:
            pass
    # 顺手清掉本工具留下的旧临时文档（只删自己命名的，不碰别人的文件）
    try:
        now = time.time()
        for f in os.listdir(WORD_WORK):
            if not f.startswith("md2docx_"):
                continue
            p = os.path.join(WORD_WORK, f)
            if os.path.isfile(p) and now - os.path.getmtime(p) > 3600:
                os.remove(p)
    except OSError:
        pass

    # ---------- 关键页 / 联系表：只给一张图，避免逐页截图 ----------
    pages_arg = None
    if "--pages" in args:
        i = args.index("--pages")
        if i + 1 < len(args):
            pages_arg = args[i + 1]
    want_sheet = "--sheet" in flags or pages_arg is not None
    if want_sheet:
        cols = 2
        if "--cols" in args:
            i = args.index("--cols")
            if i + 1 < len(args):
                cols = int(args[i + 1])
        try:
            total = int(subprocess.run(["pdfinfo", final], capture_output=True,
                                       text=True).stdout.split("Pages:")[1].split()[0])
        except (IndexError, ValueError, FileNotFoundError, TypeError):
            total = 1
        if pages_arg:
            pages = [int(x) for x in re.split(r"[,\s]+", pages_arg) if x.strip().isdigit()]
        else:
            pages = pick_key_pages(final, total)
        imgs = render_pages(final, pages)
        sheet_path = os.path.splitext(final)[0] + "_联系表.png"
        got = make_sheet(imgs, sheet_path, cols=cols)
        if got:
            print("联系表：%s（第 %s 页）" % (got, ",".join(str(p) for p, _ in imgs)))
            if "--no-open" not in flags:
                subprocess.run(["open", got])
            return 0
        print("拼图失败（需要 Pillow：pip install pillow），已导出整份 PDF：%s" % final,
              file=sys.stderr)

    print("预览：%s" % final)
    if n0:
        print("提示：Word 里留着 1 个临时文档窗口，自行关掉即可（脚本已尝试自动关）。")
    if "--no-open" not in flags:
        subprocess.run(["open", final])
    return 0


# ============================================================
# ⑤ 一键：转换 + 验收 + 摘要
# ============================================================

def cmd_make(argv):
    argv = list(argv)
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
                profile = _profile_path_or_die(val)
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
        cmd = [sys.executable, __file__, "batch", md,
               "--profiles", ",".join(profiles), "--outdir", d]
        r = run(cmd)
        if r.returncode != 0:
            print(r.stdout[-800:] + r.stderr[-800:], file=sys.stderr)
            die("批量转换失败")
        if not quiet:
            print(r.stdout.rstrip())
        for p in profiles:
            path = _resolve_or_die(p)
            label = os.path.splitext(os.path.basename(path))[0]
            fp = os.path.join(d, "%s_%s.docx" % (base, label))
            if os.path.exists(fp):
                made.append((fp, path))
    else:
        if not out:
            out = os.path.splitext(md)[0] + ".docx"
        # 显式指定档案：这样不依赖任何侧车文件也能复现
        if not profile:
            profile = _profile_path_or_die("technical")
        cmd = [sys.executable, __file__, "convert", md, out,
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
        r = run([sys.executable, __file__, "verify", path])
        good = r.returncode == 0
        ok = ok and good
        report.append((path, prof, good, r.stdout))

    # ---------- 4. 摘要 ----------
    print()
    for path, prof, good, detail in report:
        size = os.path.getsize(path) // 1024 if os.path.exists(path) else 0
        print("%s %s（%d KB）" % ("✓" if good else "✗", path, size))
        if prof:
            r = run([sys.executable, __file__, "configure", "--report", prof])
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
            r = run([sys.executable, __file__, "preview", path] + extra)
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


# ============================================================
# ⑥ 排查：逐段打印实际样式
# ============================================================


def cmd_styles(argv):
    """打印每段的样式 / 字体 / 字号 / 缩进，用来定位「这段格式为什么不对」"""
    if not argv or argv[0] in ("-h", "--help"):
        print("用法：python3 md2docx.py styles 报告.docx")
        return 0
    docx = argv[0]
    if not os.path.isfile(docx):
        die("找不到文件：%s" % docx)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    ns = {"w": W[1:-1]}
    z = zipfile.ZipFile(docx)
    sty = {}
    root = ET.fromstring(z.read("word/styles.xml"))
    for s in root.findall("w:style", ns):
        sid = s.get(W + "styleId")
        ind = s.find("w:pPr/w:ind", ns)
        rp = s.find("w:rPr", ns)
        f = rp.find("w:rFonts", ns) if rp is not None else None
        sz = rp.find("w:sz", ns) if rp is not None else None
        sty[sid] = dict(ind=(ind.get(W + "firstLineChars") if ind is not None else None),
                        ea=(f.get(W + "eastAsia") if f is not None else None),
                        sz=(int(sz.get(W + "val")) / 2 if sz is not None else None))
    body = ET.fromstring(z.read("word/document.xml")).find("w:body", ns)
    print("%-16s %-9s %-6s %-11s %s" % ("样式", "字体", "字号", "缩进(字符)", "文本"))
    print("-" * 84)
    for el in body:
        if el.tag != W + "p":
            continue
        ps = el.find("w:pPr/w:pStyle", ns)
        sid = ps.get(W + "val") if ps is not None else "Normal"
        d = el.find("w:pPr/w:ind", ns)
        direct = d.get(W + "firstLineChars") if d is not None else None
        info = sty.get(sid, {})
        ch = direct if direct else info.get("ind")
        txt = "".join(t.text or "" for t in el.iter(W + "t"))
        print("%-16s %-9s %-6s %-11s %s" % (sid, info.get("ea") or "", info.get("sz") or "",
                                            ch or "", txt[:40]))
    return 0


# ============================================================
# 入口
# ============================================================

COMMANDS = {
    "make": "cmd_make",
    "batch": "cmd_batch",
    "verify": "cmd_verify",
    "preview": "cmd_preview",
    "styles": "cmd_styles",
    "convert": "cmd_convert",
}

USAGE = """md2docx —— Markdown 报告转 Word

  python3 md2docx.py 报告.md [输出.docx] [--profile 档案]    转换（默认）
  python3 md2docx.py make 报告.md [...]                      转换 + 验收 + 摘要
  python3 md2docx.py batch 报告.md --profiles a,b,c          一份 md 出多版
  python3 md2docx.py verify 报告.docx [--json]               交付验收
  python3 md2docx.py preview 报告.docx [--sheet]             导 PDF / 联系表
  python3 md2docx.py styles 报告.docx                        逐段看样式
  python3 md2docx.py --list-profiles                         看有哪些档案

格式档案用 configure.py 生成或修改：python3 configure.py --list
"""


def _resolve_or_die(arg):
    """档案名/路径 -> (名字, 绝对路径)，找不到就报错退出"""
    name, path = resolve_profile(arg)
    if arg and not path:
        die("找不到 profile：%s（用 --list-profiles 看看）" % arg)
    return name, path


def _profile_path_or_die(arg):
    """只要路径的场合（make/batch 里传参用）"""
    return _resolve_or_die(arg)[1]


def main(argv):
    if not argv:
        print(USAGE)
        return 0
    if argv[0] in ("-h", "--help"):
        print(USAGE)
        return 0
    if "--list-profiles" in argv or argv[0] == "profiles":
        list_profiles()
        return 0
    if argv[0] == "configure":
        # 转发到 configure.py，少记一个入口
        cfg = os.path.join(HERE, "configure.py")
        os.execv(sys.executable, [sys.executable, cfg] + argv[1:])
    cmd = argv[0] if argv[0] in COMMANDS else "convert"
    rest = argv[1:] if cmd != "convert" else argv
    if cmd == "convert" and argv[0] in ("convert",):
        rest = argv[1:]
    return globals()[COMMANDS[cmd]](rest)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]) or 0)
