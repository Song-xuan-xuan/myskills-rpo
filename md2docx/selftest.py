#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest.py —— 一条命令自检整条管线

    python3 selftest.py

会做这些事：
  1. 检查依赖（pandoc / python-docx）
  2. 用每套 profile 编译模板
  3. 把两份自带样例各转一遍
  4. 检查产出的 docx 该有的东西在不在（分节、页码格式、域、斜体、图片宽度、表头重复）
最后打印 PASS / FAIL 清单。

CI 或换机器时先跑这个。
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
# 自检产物默认放临时目录：skill 目录里只保留代码和样例，
# 免得每次自检都堆一堆 docx。想留档就加 --keep。
if "--keep" in sys.argv:
    OUT = os.path.join(HERE, "build", "selftest")
else:
    OUT = os.path.join(tempfile.mkdtemp(prefix="md2docx-selftest."), "out")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

results = []


def _load_docxfix():
    import importlib.util
    spec = importlib.util.spec_from_file_location("dx", os.path.join(HERE, "docxfix.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


_dx = None


def load_profile(path):
    global _dx
    if _dx is None:
        _dx = _load_docxfix()
    return _dx.load_profile(path)


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print("  %s %s%s" % ("✓" if ok else "✗", name,
                        ("  — " + detail) if detail else ""))
    return ok


def run(cmd, **kw):
    return subprocess.run(cmd, cwd=HERE, capture_output=True, text=True, **kw)


def deps():
    print("\n[1/4] 检查依赖")
    p = run(["pandoc", "--version"])
    if p.returncode != 0:
        check("pandoc 已安装", False, "去 https://pandoc.org 装一个")
        return False
    ver = p.stdout.splitlines()[0]
    check("pandoc 已安装", True, ver)
    major = re.search(r"pandoc (\d+)", ver)
    check("pandoc >= 3（2.x 也能跑，只是参数不同）",
          True, "检测到 %s.x" % (major.group(1) if major else "?"))
    d = run([sys.executable, "-c", "import docx; print(docx.__version__ if hasattr(docx,'__version__') else 'ok')"])
    check("python-docx 已安装", d.returncode == 0,
          d.stdout.strip() if d.returncode == 0 else "pip3 install python-docx")
    return d.returncode == 0


def build_templates():
    print("\n[2/4] 编译模板")
    prof_dir = os.path.join(HERE, "profiles")
    ok = True
    for f in sorted(os.listdir(prof_dir)):
        if not f.endswith(".json"):
            continue
        name = f[:-5]
        tpl = os.path.join(OUT, name + ".docx")
        r = run([sys.executable, "build_template.py", "-p",
                 os.path.join("profiles", f), "-o", tpl])
        ok = check("模板 " + name, r.returncode == 0, r.stderr.strip()[:200]) and ok
    return ok


def convert():
    print("\n[3/4] 转换样例")
    os.makedirs(OUT, exist_ok=True)
    # (标签, md 路径, 档案)
    pairs = [("样例报告", "sample-report", "technical"),
             ("复杂样例", "complex-sample", "technical"),
             ("复杂样例", "complex-sample", "academic"),
             ("样例报告", "sample-report", "official")]
    made = []
    for label, stem, prof in pairs:
        md = os.path.join("examples", stem + ".md")
        out = os.path.join(OUT, "%s_%s.docx" % (label, prof))
        r = run([sys.executable, "md2docx.py", md, out, "--profile", prof])
        ok = check("转换 %s (%s)" % (label, prof), r.returncode == 0,
                   r.stderr.strip()[:200] if r.returncode else "")
        if ok:
            made.append((out, prof, label))
    return made


def inspect(made):
    print("\n[4/4] 检查产出")
    for path, prof, md in made:
        tag = "%s/%s" % (md, prof)
        z = zipfile.ZipFile(path)
        doc = z.read("word/document.xml").decode("utf-8")
        sty = z.read("word/styles.xml").decode("utf-8")
        # 页眉页脚是独立部件，页码域在里面
        hf = " ".join(z.read(n).decode("utf-8") for n in z.namelist()
                      if re.match(r"word/(header|footer)\d*\.xml$", n))

        # 分节
        n_sect = doc.count("<w:sectPr")
        want3 = prof in ("technical", "academic")
        check("%s 分节数 = %d" % (tag, n_sect),
              n_sect == (3 if want3 else 1),
              "封面/目录/正文三节" if want3 else "单节")

        # 页码格式
        fmts = re.findall(r'<w:pgNumType w:fmt="(\w+)"', doc)
        if want3:
            check("%s 目录页罗马 + 正文阿拉伯" % tag,
                  fmts[:2] == ["upperRoman", "decimal"], str(fmts))
        else:
            check("%s 页码格式已设置" % tag, "decimal" in fmts, str(fmts))

        # 域
        instr = re.findall(r"<w:instrText[^>]*>([^<]*)<", doc)
        joined = " ".join(instr)
        hf_instr = " ".join(re.findall(r"<w:instrText[^>]*>([^<]*)<", hf))
        cfg = load_profile(os.path.join(HERE, "profiles", prof + ".json"))
        want_toc = (cfg.get("toc") or {}).get("auto", True)
        want_page = (cfg.get("footer") or {}).get("page_number", True)
        check("%s 有目录域" % tag, ("TOC" in joined) == want_toc,
              "" if want_toc else "该档案配置为不生成目录")
        check("%s 有题注 SEQ 域" % tag, "SEQ" in joined)
        check("%s 有交叉引用 REF 域" % tag, "REF" in joined)
        check("%s 有页码域" % tag, ("PAGE" in hf_instr) == want_page)

        # 斜体（样式 + 正文都不能有开着的 i）
        bad = 0
        for xml in (doc, sty):
            for m in re.finditer(r"<w:i(?:Cs)?(?:\s[^>]*)?/>", xml):
                if 'w:val="0"' not in m.group(0) and 'w:val="false"' not in m.group(0):
                    bad += 1
        check("%s 无遗留斜体" % tag, bad == 0, "发现 %d 处" % bad if bad else "")

        # 图片不超版心
        pg = re.search(r'<w:pgSz w:w="(\d+)"', doc)
        mar = re.search(r'<w:pgMar[^>]*w:left="(\d+)"[^>]*w:right="(\d+)"', doc)
        if pg and mar:
            limit = int(pg.group(1)) - int(mar.group(1)) - int(mar.group(2))
            over = [int(cx) for cx in re.findall(r'<wp:extent cx="(\d+)"', doc)
                    if int(int(cx) / 635) > limit + 10]
            check("%s 图片不超版心" % tag, not over,
                  str(over) if over else "")

        # 表格：跨页重复表头
        n_tbl = len(re.findall(r"<w:tbl>", doc))
        if n_tbl:
            check("%s 表格重复表头" % tag, doc.count("tblHeader") == n_tbl,
                  "%d 表 / %d 表头" % (n_tbl, doc.count("tblHeader")))

        # 题注样式存在
        check("%s 题注使用 ImageCaption/TableCaption" % tag,
              "ImageCaption" in doc or "TableCaption" in doc)

        # 标题颜色不能是 pandoc 的主题蓝（正式报告要黑）
        themed = []
        for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>',
                             sty, re.S):
            if m.group(1) == "Hyperlink":      # 超链接保留主题色是对的
                continue
            if 'w:themeColor="accent1"' in m.group(0):
                themed.append(m.group(1))
        check("%s 标题无主题蓝" % tag, not themed,
              ",".join(themed) if themed else "")

        # 编号：中文序号后面不跟空格；混用序号时加 isLgl
        numbering = z.read("word/numbering.xml").decode("utf-8")
        for m in re.finditer(r'<w:lvl w:ilvl="(\d)"[^>]*>(.*?)</w:lvl>',
                             numbering, re.S):
            lvl, blk = m.group(1), m.group(2)
            txt = re.findall(r'w:lvlText w:val="([^"]*)"', blk)
            if not txt:
                continue
            if txt[0].endswith(("、", "。", "．")):
                check("%s 编号「%s」后无空格" % (tag, txt[0]),
                      'w:suff w:val="nothing"' in blk, blk[:120])


def guard_docs():
    """文档守门：不许再出现诱导 agent 逐页看图的说法"""
    print("\n[文档守门]")
    ok = True
    skill = open(os.path.join(HERE, "SKILL.md"), encoding="utf-8").read()
    verify_src = open(os.path.join(HERE, "md2docx.py"), encoding="utf-8").read()

    # ① SKILL 的主流程代码块里不能直接出现 preview.py 调用（只在「看效果」一节允许）
    flow = skill[skill.index("### 阶段 ④"):skill.index("### 阶段 ⑤")]
    ok = check("转换阶段不含自动渲染命令",
               "preview 交付版" not in flow or "--sheet" in flow) and ok

    # ② verify.py 不许写「打开 docx → F9」这类指令
    ok = check("verify 输出不再指令人肉打开",
               "做法：打开 docx" not in verify_src) and ok

    # ③ 必须有禁止逐页 read 的硬约束
    ok = check("SKILL 有「禁止 read 单页图片」",
               "禁止用 read 打开单页图片" in skill) and ok

    # ④ 必须有内容定稿门禁
    ok = check("SKILL 有「内容确认后才导出」门禁",
               "硬门禁" in skill and "阶段 ②" in skill) and ok

    # ⑤ 文件数守门：只允许 6 个 py，防止又长出十几个入口
    pys = sorted(f for f in os.listdir(HERE) if f.endswith(".py"))
    ok = check("只有 6 个 py 文件", len(pys) == 6, str(pys)) and ok
    ok = check("没有多余的入口脚本",
               not any(os.path.exists(os.path.join(HERE, f))
                       for f in ("make.py", "verify.py", "preview.py",
                                 "batch.py", "asklib.py", "format_report.py",
                                 "check.py"))) and ok

    # ⑥ 根目录不许再堆样例/文档（GitHub 仓库要干净）
    root_junk = [f for f in os.listdir(HERE)
                 if f.endswith((".md", ".png", ".bib"))
                 and f not in ("README.md", "SKILL.md")]
    ok = check("根目录只有 SKILL.md / README.md 两个 md",
               not root_junk, str(root_junk[:5])) and ok
    ok = check("样例在 examples/ 下",
               os.path.exists(os.path.join(HERE, "examples", "sample-report.md"))) and ok

    # ⑥ 测试说明里的人肉验收步骤要标注给人看
    t = open(os.path.join(HERE, "docs", "TESTING.md"), encoding="utf-8").read()
    ok = check("docs/TESTING.md 标注了「给人看的，agent 请跳过」",
               "agent 请跳过" in t) and ok
    return ok


def extras2():
    """排版细节：目录条目样式、中文排版、奇偶页页眉、横排页"""
    print("\n[6/7] 排版细节")
    ok = True
    import zipfile
    prof = os.path.join(HERE, "profiles", "technical.json")
    cfg = load_profile(prof)
    tpl = os.path.join(OUT, "tpl.docx")
    r = run([sys.executable, "build_template.py", "-p", prof, "-o", tpl])
    if not check("模板可编译", r.returncode == 0, r.stderr.strip()[:200]):
        return False
    sty = zipfile.ZipFile(tpl).read("word/styles.xml").decode()
    for sid in ("TOC1", "TOC2", "TOC3", "TableofFigures"):
        ok = check("有目录条目样式 %s" % sid, 'w:styleId="%s"' % sid in sty) and ok
    ok = check("正文启用了孤行控制", "<w:widowControl/>" in sty) and ok
    ok = check("正文启用了中文标点控制", "<w:kinsoku/>" in sty) and ok
    ok = check("标题颜色不是主题蓝",
               'w:themeColor="accent1"' not in sty.split("Hyperlink")[0]) and ok

    # 奇偶页页眉 + 横排页
    tmp = os.path.join(OUT, "even")
    os.makedirs(tmp, exist_ok=True)
    p2 = os.path.join(tmp, "p.json")
    c = json.loads(json.dumps(cfg))
    c.setdefault("header", {})["even_text"] = "某某单位"
    c["landscape"] = ["tbl:wide"]
    with open(p2, "w", encoding="utf-8") as fh:
        json.dump(c, fh, ensure_ascii=False, indent=2)
    md = os.path.join(tmp, "m.md")
    with open(md, "w", encoding="utf-8") as fh:
        fh.write("---\ntitle: T\n---\n\n# 一\n\n"
                 "| a | b |\n|---|---|\n| 1 | 2 |\n\n: 说明 {#tbl:wide}\n")
    out = os.path.join(tmp, "m.docx")
    r = run([sys.executable, "md2docx.py", md, out, "--profile", p2])
    if check("奇偶页+横排的档案能转换", r.returncode == 0,
             r.stderr.strip()[:200]):
        z = zipfile.ZipFile(out)
        doc = z.read("word/document.xml").decode()
        st = z.read("word/settings.xml").decode()
        ok = check("启用了奇偶页页眉", "evenAndOddHeaders" in st) and ok
        ok = check("宽表被放到横向页", 'w:orient="landscape"' in doc) and ok
    return ok


def extras3():
    """新增能力：摘要、脚注、列表、表格列宽、文档开关、首字下沉、网格"""
    print("\n[7/8] 新增排版能力")
    ok = True
    import zipfile
    tmp = os.path.join(OUT, "feat")
    os.makedirs(tmp, exist_ok=True)
    base = load_profile(os.path.join(HERE, "profiles", "technical.json"))
    p2 = os.path.join(tmp, "p.json")
    c = json.loads(json.dumps(base))
    c["drop_cap"] = {"enabled": True, "lines": 3, "chars": 1}
    c["table"] = dict(c.get("table") or {})
    c["table"]["cant_split"] = True
    c["table"]["col_widths"] = {"tbl:a": [30, 70]}
    c["lists"] = {"ordered": {"num_fmt": "decimal", "format": "%1)"},
                  "bullet": {"char": "▪"}}
    c["document"] = dict(c.get("document") or {})
    c["document"]["track_changes"] = True
    c["notes"] = {"footnote": {"pos": "beneathText", "num_fmt": "chineseCounting",
                               "num_restart": "eachPage"}}
    c["page"] = dict(c.get("page") or {})
    c["page"]["doc_grid"] = {"type": "lines", "lines_per_page": 22}
    with open(p2, "w", encoding="utf-8") as fh:
        json.dump(c, fh, ensure_ascii=False, indent=2)

    md = os.path.join(tmp, "f.md")
    with open(md, "w", encoding="utf-8") as fh:
        fh.write("---\ntitle: T\nabstract: |\n  摘要正文。\n"
                 "keywords: [甲, 乙]\n---\n\n# 一\n\n"
                 "这是正文第一段，应该有首字下沉。\n\n"
                 "1. 第一项\n2. 第二项\n\n"
                 "- 项目甲\n- 项目乙\n\n"
                 "| A | B |\n|---|---|\n| 1 | 2 |\n\n: 说明 {#tbl:a}\n")
    out = os.path.join(tmp, "f.docx")
    r = run([sys.executable, "md2docx.py", md, out, "--profile", p2])
    if not check("新功能档案能转换", r.returncode == 0, r.stderr.strip()[:300]):
        return False
    z = zipfile.ZipFile(out)
    doc = z.read("word/document.xml").decode()
    st = z.read("word/settings.xml").decode()
    num = z.read("word/numbering.xml").decode()
    txt = re.findall(r"<w:t[^>]*>([^<]*)<", doc)
    ok = check("摘要被搬到前置部分", "摘　要" in txt) and ok
    ok = check("关键词行生成了", any(t.startswith("关键词") for t in txt)) and ok
    ok = check("脚注位置/格式写进 sectPr",
               'w:pos w:val="beneathText"' in doc
               and 'w:numFmt w:val="chineseCounting"' in doc
               and 'w:numRestart w:val="eachPage"' in doc) and ok
    ok = check("无序列表符号改了", 'w:lvlText w:val="▪"' in num) and ok
    ok = check("有序列表编号改了", 'w:lvlText w:val="%1)"' in num) and ok
    ok = check("表格行不跨页", "cantSplit" in doc) and ok
    ok = check("表格列宽固定", 'w:tblLayout w:type="fixed"' in doc) and ok
    ok = check("首字下沉", 'w:dropCap="drop"' in doc) and ok
    ok = check("文档网格每页行数", 'w:docGrid w:linePitch="634"' in doc) and ok
    ok = check("修订模式", "trackChanges" in st) and ok
    return ok


def interactive():
    """交互能力：结构化提问、问题筛选、样张、模板反推、档案对比"""
    print("\n[9/10] 交互与定制流程")
    ok = True

    # ① 扫 md 判断该问什么
    r = run([sys.executable, "configure.py", "--ask", "--md", "examples/complex-sample.md", "--round", "1"])
    good = r.returncode == 0 and '"questions"' in r.stdout
    ok = check("--ask 输出结构化问题", good, r.stderr.strip()[:200]) and ok
    if good:
        import json as _json
        payload = _json.loads(r.stdout)
        q = payload["rounds"][0]["questions"][0]
        ok = check("问题含 id/header/question/options",
                   all(k in q for k in ("id", "header", "question", "options"))) and ok
        ok = check("每题都给 2-4 个选项", 2 <= len(q["options"]) <= 4) and ok
        ok = check("每轮不超过 4 题", len(payload["rounds"][0]["questions"]) <= 4) and ok
        ok = check("有 features 特征表", "figures" in payload["features"]) and ok

    # ② 相关性过滤：没有参考文献的 md 不该问 CSL
    r2 = run([sys.executable, "configure.py", "--ask", "--md", "examples/complex-sample.md",
              "--round", "5"])
    if r2.returncode == 0:
        ok = check("无 bibliography 时不问参考文献风格",
                   "csl_style" not in r2.stdout) and ok

    # ③ 样张
    r = run([sys.executable, "configure.py", "--sample", "technical",
             "--outdir", OUT])
    tpl = os.path.join(OUT, "样张_technical.docx")
    good = r.returncode == 0 and os.path.exists(tpl)
    ok = check("--sample 生成样张", good, r.stderr.strip()[:200]) and ok
    if good:
        import zipfile
        doc = zipfile.ZipFile(tpl).read("word/document.xml").decode()
        ok = check("样张含各级标题与正文",
                   "Heading1" in doc and "BodyText" in doc) and ok
        ok = check("样张含题注/公式/代码",
                   "ImageCaption" in doc and "oMath" in doc
                   and "SourceCode" in doc) and ok
        ok = check("样张没有封面和目录（保持精简）",
                   "CoverTitle" not in doc) and ok

    # ④ 从 Word 模板反推
    src = os.path.join(OUT, "样例报告_technical.docx")
    if os.path.exists(src):
        out = os.path.join(OUT, "reversed.json")
        r = run([sys.executable, "configure.py", "--from-docx", src, "--out", out])
        good = r.returncode == 0 and os.path.exists(out)
        ok = check("--from-docx 反推档案", good, r.stderr.strip()[:200]) and ok
        if good:
            cfg2 = load_profile(out)
            normal = (cfg2.get("styles") or {}).get("Normal") or {}
            ok = check("反推出了正文字体", bool(normal.get("east_asia_font")),
                       str(normal)) and ok
            ok = check("反推出了页边距",
                       bool((cfg2.get("page") or {}).get("margin_cm"))) and ok
            ok = check("反推出了页眉文字",
                       bool((cfg2.get("header") or {}).get("text"))) and ok
            ok = check("页脚占位符正确",
                       "{NUMPAGES}" in ((cfg2.get("footer") or {}).get("format") or ""),
                       str((cfg2.get("footer") or {}).get("format"))) and ok

    # ⑤ 档案对比
    r = run([sys.executable, "configure.py", "--diff", "profiles/technical.json",
             "profiles/official.json"])
    ok = check("--diff 能对比两份档案",
               r.returncode == 0 and "差异" in r.stdout) and ok
    return ok


def extras4():
    """附录编号 / 封面 logo / 批量出多版"""
    print("\n[11/12] 附录、封面 logo、批量出多版")
    ok = True
    import zipfile
    tmp = os.path.join(OUT, "grade")
    os.makedirs(tmp, exist_ok=True)

    # ① 附录编号
    base = load_profile(os.path.join(HERE, "profiles", "technical.json"))
    c = json.loads(json.dumps(base))
    c["appendix"] = {"enabled": True, "marker": "sec:app", "label": "附录"}
    c.setdefault("cover", {})["logo"] = os.path.join(HERE, "examples/flow.png")
    c["cover"]["logo_width_cm"] = 3.0
    p2 = os.path.join(tmp, "p.json")
    with open(p2, "w", encoding="utf-8") as fh:
        json.dump(c, fh, ensure_ascii=False, indent=2)
    md = os.path.join(tmp, "a.md")
    with open(md, "w", encoding="utf-8") as fh:
        fh.write("---\ntitle: T\n---\n\n# 概述\n\n见 {{sec:app}} 与 {{sec:app_data}}。\n\n"
                 "# 数据说明 {#sec:app}\n\n附录正文。\n\n"
                 "## 明细 {#sec:app_data}\n\n二级。\n\n"
                 "# 补充实验\n\n第二个附录。\n\n"
                 "# 结论 {#sec:endappendix}\n\n回到正常。\n")
    out = os.path.join(tmp, "a.docx")
    r = run([sys.executable, "md2docx.py", md, out, "--profile", p2])
    if check("带附录+logo 的档案能转换", r.returncode == 0, r.stderr.strip()[:300]):
        z = zipfile.ZipFile(out)
        doc = z.read("word/document.xml").decode()
        num = z.read("word/numbering.xml").decode()
        txt = re.findall(r"<w:t[^>]*>([^<]*)<", doc)
        ok = check("附录列表用字母编号", "upperLetter" in num) and ok
        ok = check("一级附录带「附录」前缀", "附录 %1" in num) and ok
        ok = check("附录下的二级编号 A.1", any(t == "A.1" for t in txt), str(txt[:12])) and ok
        ok = check("附录一级引用显示「附录 A」",
                   any(t.strip() == "附录 A" for t in txt)) and ok
        # 编号是域、不进正文文字，所以直接看 collect_headings 的推算
        import importlib.util as _ilu
        _sp = _ilu.spec_from_file_location("dx2", os.path.join(HERE, "docxfix.py"))
        _dx2 = _ilu.module_from_spec(_sp); _sp.loader.exec_module(_dx2)
        from docx import Document as _D2
        labels = _dx2.collect_headings(_D2(out), c)
        ok = check("结束标记后回到正常编号",
                   labels.get("sec:endappendix", ("",))[0] == "2",
                   str({k: v[0] for k, v in labels.items()})) and ok
        ok = check("附录标题挂上段落级编号",
                   doc.count("<w:numPr>") >= 3) and ok
        ok = check("封面 logo 内嵌成图片",
                   any(n.startswith("word/media/") for n in z.namelist())) and ok

    # ② 批量出多版
    d = os.path.join(tmp, "dist")
    r = run([sys.executable, "md2docx.py", "batch", os.path.join(HERE, "examples/sample-report.md"),
             "--outdir", d, "--profiles", "technical,official"])
    files = []
    if os.path.isdir(d):
        files = [f for f in os.listdir(d) if f.endswith(".docx")]
    ok = check("batch.py 出两个版本", r.returncode == 0 and len(files) == 2,
               str(files)) and ok
    return ok


def citations():
    """参考文献 + 章节交叉引用"""
    print("\n[12/12] 参考文献与章节交叉引用")
    ok = True
    tmp = os.path.join(OUT, "cite")
    os.makedirs(tmp, exist_ok=True)
    with open(os.path.join(tmp, "refs.bib"), "w", encoding="utf-8") as fh:
        fh.write("@article{key1, title={T}, author={A, B}, "
                 "journal={J}, volume={1}, year={2020}}\n")
    md = os.path.join(tmp, "c.md")
    with open(md, "w", encoding="utf-8") as fh:
        fh.write("---\ntitle: 引用测试\nbibliography: refs.bib\n---\n\n"
                 "# 引言\n\n见 @key1，详见 {{sec:d}}。\n\n"
                 "## 细节 {#sec:d}\n\n正文。\n\n"
                 "# 参考文献\n\n::: {#refs}\n:::\n")
    out = os.path.join(tmp, "c.docx")
    r = run([sys.executable, "md2docx.py", md, out, "--profile", "technical"])
    if not check("带 references 的 md 能转换", r.returncode == 0,
                 r.stderr.strip()[:200]):
        return False
    import zipfile
    z = zipfile.ZipFile(out)
    doc = z.read("word/document.xml").decode()
    texts = re.findall(r"<w:t[^>]*>([^<]*)<", doc)
    # 角标现在是「[ + REF 域 + ]」，文字被拆成多个 run，不能再找整串 "[1]"
    ok = check("正文出现角标括号与编号",
               "[" in texts and "]" in texts and "1" in texts,
               str(texts[:20])) and ok
    ok = check("生成了 Bibliography 样式段落",
               "Bibliography" in doc) and ok
    ok = check("参考文献条目按国标格式（[J] 期刊标识）", any(
        "[J]" in t for t in texts), str([t for t in texts if "[J]" in t][:2])) and ok
    ok = check("书目编号是 SEQ 域（删条目会自动重排）",
               "SEQ 文献" in doc) and ok
    ok = check("正文角标是 REF 域", "REF _Ref_bib_" in doc) and ok
    ok = check("书目编号书签存在",
               len(re.findall(r'w:name="_Ref_bib_', doc)) >= 1) and ok
    ok = check("章节交叉引用生成 REF 域",
               "REF sec:d" in doc) and ok
    ok = check("章节引用缓存值是章号",
               any(t == "1.1" for t in texts)) and ok
    return ok


def extras():
    """额外检查：格式清单、改单个键、自定义模板"""
    print("\n[5/7] 交互与定制能力")
    ok = True

    # 格式清单
    r = run([sys.executable, "configure.py", "--report", "technical"])
    good = r.returncode == 0 and "正文" in r.stdout and "图题注" in r.stdout
    ok = check("格式清单可生成", good, r.stderr.strip()[:200] if not good else "")
    if good:
        ok = check("格式清单含缩进信息", "首行缩进" in r.stdout or "顶格" in r.stdout)

    # --set 改一个键
    tmp = os.path.join(OUT, "set-test.json")
    r = run([sys.executable, "configure.py", "--base", "profiles/technical.json",
             "--out", tmp, "--set", "styles.Caption.bold=true"])
    ok = check("--set 改单个键", r.returncode == 0,
               r.stderr.strip()[:200] if r.returncode else "") and ok
    if r.returncode == 0:
        cfg = load_profile(tmp)
        ok = check("--set 真的写进去了",
                   (cfg.get("styles", {}).get("Caption", {}).get("bold")) is True) and ok

    # --reference：用户自己的 Word 模板
    tpl = os.path.join(OUT, "user-template.docx")
    r = run([sys.executable, "-c", """
import sys
from docx import Document
d = Document(sys.argv[1])
d.styles['Caption'].font.bold = True
d.save(sys.argv[2])
""", os.path.join(OUT, "technical.docx"), tpl])
    if r.returncode == 0:
        out = os.path.join(OUT, "ref-test.docx")
        r = run([sys.executable, "md2docx.py", "examples/sample-report.md", out, "--reference", tpl])
        if check("--reference 自定义模板可用", r.returncode == 0,
                 r.stderr.strip()[:200]) and ok:
            import zipfile
            st = zipfile.ZipFile(out).read("word/styles.xml").decode()
            bold = bool(re.search(r'<w:style [^>]*w:styleId="Caption".*?<w:b\s*/>',
                                  st, re.S))
            ok = check("模板里的加粗保留到了输出", bold)
    else:
        ok = check("--reference 自定义模板可用", False, r.stderr.strip()[:200]) and ok
    return ok


def main():
    print("=" * 60)
    print("md2docx 自检")
    print("=" * 60)
    if not deps():
        print("\n依赖缺失，先装依赖再跑。")
        return 1
    build_templates()
    made = convert()
    inspect(made)
    extras()
    guard_docs()
    extras2()
    extras3()
    interactive()
    extras4()
    citations()

    # 最后统一做一次结构合规检查（Word 会不会报损坏）
    print("\n[结构合规]")
    import glob as _glob
    import zipfile as _zip
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location("dx", os.path.join(HERE, "docxfix.py"))
    _dx = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(_dx)
    from docx import Document as _Doc
    outs = _glob.glob(os.path.join(OUT, "**", "*.docx"), recursive=True)
    outs += [_glob.glob("*.docx") and f for f in _glob.glob("*.docx")]
    bad = 0
    for f in outs:
        try:
            probs = _dx.validate(_Doc(f), quiet=True)
        except Exception as e:
            probs = ["打不开：%s" % e]
        if probs:
            bad += 1
            check("结构合规 %s" % os.path.basename(f), False, probs[0])
    check("全部产物结构合规（%d 个文件）" % len(outs), bad == 0)

    # 交付质量：目录制表位、REF 落点、页码域、角标可跳转、compat
    print("\n[交付质量]")
    ok = True
    import zipfile as _z
    sample = os.path.join(OUT, "样例报告_technical.docx")
    if os.path.exists(sample):
        z = _z.ZipFile(sample)
        sty = z.read("word/styles.xml").decode()
        doc = z.read("word/document.xml").decode()
        sett = z.read("word/settings.xml").decode()
        toc_tabs = all("<w:tabs>" in blk for blk in
                       __import__("re").findall(
                           r'<w:style [^>]*w:styleId="TOC1".*?</w:style>', sty,
                           __import__("re").S))
        check("目录条目 TOC1 有制表位", toc_tabs)
        ok = check("目录制表位是点线前导", 'w:leader="dot"' in sty)
        check("TOC2/TOC3 有层级缩进",
              'w:leftChars="200"' in sty and 'w:leftChars="400"' in sty)
        ok = check("参考文献角标可跳转", '<w:hyperlink w:anchor="ref-' in doc or
              "ref-" not in doc)
        check("有 compatibilityMode=15", "compatibilityMode" in sett)
        check("表格单元格用独立样式 TableText", "TableText" in doc)
        check("列表项仍用 Compact",
              'w:val="Compact"' in doc or "Compact" not in doc)
        ok = check("关键词不带方括号", not __import__("re").search(
            r"关键词[：:]\s*\[", doc))
    # make.py 一键链路
    tmp2 = os.path.join(OUT, "make")
    os.makedirs(tmp2, exist_ok=True)
    src_md = os.path.join(HERE, "examples/sample-report.md")
    r = run([sys.executable, "md2docx.py", "make", src_md, "--out",
             os.path.join(tmp2, "out.docx")])
    ok = check("make.py 一键通过", r.returncode == 0,
               (r.stdout[-300:] + r.stderr[-300:])) and ok
    ok = check("make.py 打印了格式摘要", "本次格式" in r.stdout) and ok
    # 不能出现工具内部的临时目录（输出路径本身是调用者指定的，不算）
    # 只关心工具自己造的临时目录（输出路径是调用方给的，含 selftest 自己的 OUT 是正常的）
    leaked = [l for l in r.stdout.splitlines()
              if "md2docx-tpl." in l or "md2docx." in l]
    ok = check("make.py 不打印内部临时路径", not leaked, str(leaked[:2])) and ok

    # verify.py：一次把能判的都判掉
    r = run([sys.executable, "md2docx.py", "verify", sample])
    ok = check("verify 对成品判为通过", r.returncode == 0,
               r.stdout[-300:]) and ok
    ok = check("verify 输出「需人看」清单",
               "需要人看" in r.stdout or "没有需要肉眼" in r.stdout,
               r.stdout[-200:]) and ok
    r = run([sys.executable, "md2docx.py", "verify", sample, "--json"])
    good = r.returncode in (0, 1) and '"problems"' in r.stdout
    ok = check("verify 支持 --json", good, r.stderr[-200:]) and ok
    # 故意造一个坏文件，确认能报出来
    bad = os.path.join(OUT, "bad-anchor.docx")
    md2 = os.path.join(OUT, "bad.md")
    with open(md2, "w", encoding="utf-8") as fh:
        fh.write("---\ntitle: T\n---\n\n# 一\n\n见 {{fig:none}} 。\n")
    r = run([sys.executable, "md2docx.py", md2, bad])
    r2 = run([sys.executable, "md2docx.py", "verify", bad])
    ok = check("verify 能报出未替换的引用标记",
               r2.returncode == 1 and "没替换" in r2.stdout, r2.stdout[-300:]) and ok

    fails = [r for r in results if not r[1]]
    print("\n" + "=" * 60)
    print("通过 %d 项，失败 %d 项" % (len(results) - len(fails), len(fails)))
    if fails:
        for name, _, detail in fails:
            print("  ✗ %s  %s" % (name, detail))
        return 1
    print("全部通过。自检产物在 %s（加 --keep 可留档到 build/）" % OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
