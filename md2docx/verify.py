#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify.py —— 交付前的一次性程序化验收

    python3 verify.py 报告.docx            # 人话报告
    python3 verify.py 报告.docx --json     # 机器可读
    python3 verify.py dist/                # 整个目录

设计目的：把「能用程序判断的东西」全部判掉，剩下真正需要眼睛看的才列出来。
这样就不用来回导 PDF、逐页截图了。

分三类输出：
  [错误] 必须修，否则交付会出问题
  [提醒] 不影响打开，但可能不是你要的效果
  [需肉眼] 程序判不了的少数几项（通常 0~2 项）

退出码：有错误 → 1，否则 0。
"""
import glob
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

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


def main():
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


if __name__ == "__main__":
    sys.exit(main())
