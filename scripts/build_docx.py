#!/usr/bin/env python3
"""
build_docx.py — 把 Markdown 正文（plan.md）排成正式的商业计划书 Word 文档。

用法:
    python build_docx.py plan.md -o "项目-商业计划书-2026-11.docx" \
        --title "××商业计划书" --subtitle "一句话定位" --org "××公司" --to "呈送：××投资" \
        --date "2026年11月" --version "V1.0" --lang zh [--no-toc] [--no-cover] [--logo logo.png]

支持的 Markdown 子集:
    # 章标题 / ## 节标题 / ### 小节 / #### 四级
    段落（连续行合并）、**粗体**、*斜体*、`代码`
    - 无序列表 / 1. 有序列表（两个空格缩进为二级）
    | 表格 |（第二行为分隔行）；表格前一段以“表 ”或“Table ”开头会被当作表题
    ![图题](路径.png)  图片（宽 15cm）+ 图题
    > 引用/提示框
    <!-- pagebreak -->  分页
    [S1] [A-03] 等引用标记按原文保留；加 --strip-refs 可在对外终稿中删除这些标记（[待补充] 保留）
"""
import argparse
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONTS = {
    "zh": {"body": "宋体", "body_ascii": "Times New Roman", "head": "黑体", "head_ascii": "Arial", "body_size": 11, "indent": True},
    "en": {"body": "Calibri", "body_ascii": "Calibri", "head": "Calibri", "head_ascii": "Calibri", "body_size": 11, "indent": False},
}
TEXT = {
    "zh": {"toc": "目  录", "confidential": "保密声明", "conf_text": "本文件所含信息属于{org}的商业秘密，仅供{to}评估之用。未经书面许可，不得复制、传播或用于其他目的。本文件中的预测基于文中列明的假设，不构成对未来业绩的承诺。", "page": "第 {p} 页", "caption_table": "表", "caption_fig": "图", "to_default": "指定接收方", "update_toc": "请在 Word 中右键此处并选择“更新域”以生成目录。"},
    "en": {"toc": "Contents", "confidential": "Confidentiality notice", "conf_text": "This document contains confidential information of {org} and is provided solely for evaluation by {to}. It may not be reproduced or distributed without written consent. Projections are based on the assumptions stated herein and do not constitute a guarantee of future performance.", "page": "Page {p}", "caption_table": "Table", "caption_fig": "Figure", "to_default": "the intended recipient", "update_toc": "Right-click here in Word and choose “Update Field” to generate the table of contents."},
}
NUM_RE = re.compile(r"^[-+]?[\d,]*\.?\d+%?$")
INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*.+?\*|`.+?`)")


def set_run_font(run, name, ascii_name, size=None, bold=None, italic=None, color=None):
    run.font.name = ascii_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_name)
    rfonts.set(qn("w:hAnsi"), ascii_name)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:cs"), ascii_name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def style_fonts(doc, lang):
    f = FONTS[lang]
    for sname, size, bold, use_head in [("Normal", f["body_size"], None, False), ("Heading 1", 16, True, True), ("Heading 2", 13.5, True, True), ("Heading 3", 12, True, True), ("Heading 4", 11, True, True), ("Title", 26, True, True), ("List Bullet", f["body_size"], None, False), ("List Number", f["body_size"], None, False), ("List Bullet 2", f["body_size"], None, False), ("List Number 2", f["body_size"], None, False)]:
        try:
            st = doc.styles[sname]
        except KeyError:
            continue
        st.font.size = Pt(size)
        if bold is not None:
            st.font.bold = bold
        name = f["head"] if use_head else f["body"]
        ascii_name = f["head_ascii"] if use_head else f["body_ascii"]
        st.font.name = ascii_name
        rpr = st.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        for k in ("w:ascii", "w:hAnsi", "w:cs"):
            rfonts.set(qn(k), ascii_name)
        rfonts.set(qn("w:eastAsia"), name)
        st.font.color.rgb = RGBColor(0x0B, 0x0B, 0x0B) if use_head else RGBColor(0x1A, 0x1A, 0x1A)
        if sname.startswith("Heading"):
            st.paragraph_format.keep_with_next = True
            st.paragraph_format.space_before = Pt(18 if sname == "Heading 1" else 12)
            st.paragraph_format.space_after = Pt(8 if sname == "Heading 1" else 6)
            st.paragraph_format.line_spacing = 1.2
        if sname == "Heading 1":
            st.paragraph_format.page_break_before = True
    normal = doc.styles["Normal"]
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.space_after = Pt(6)


def add_inline(paragraph, text, lang, size=None, color=None):
    f = FONTS[lang]
    parts = INLINE_RE.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, f["body"], f["body_ascii"], size=size, bold=True, color=color)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            set_run_font(run, f["body"], f["body_ascii"], size=size, italic=True, color=color)
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            set_run_font(run, "等线", "Consolas", size=size, color=color)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, f["body"], f["body_ascii"], size=size, color=color)


def add_field(paragraph, instr):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr_el)
    run._r.append(fld_sep)
    placeholder = OxmlElement("w:t")
    placeholder.text = ""
    run._r.append(placeholder)
    run._r.append(fld_end)
    return run


def shade_paragraph(paragraph, fill="F2F2F2"):
    ppr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def shade_cell(cell, fill="D9E1F2"):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def set_cell_margins(table, top=40, bottom=40, left=80, right=80):
    tbl_pr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)


def enable_update_fields(doc):
    settings = doc.settings.element
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    settings.append(uf)


def build_cover(doc, args, lang):
    f = FONTS[lang]
    T = TEXT[lang]
    if args.logo and os.path.exists(args.logo):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.add_run().add_picture(args.logo, width=Cm(4))
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(args.title)
    set_run_font(r, f["head"], f["head_ascii"], size=28, bold=True)
    if args.subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(args.subtitle)
        set_run_font(r, f["body"], f["body_ascii"], size=14, color="52514E")
    for _ in range(6):
        doc.add_paragraph()
    for line in [args.to, args.org, args.date, args.version]:
        if line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(line)
            set_run_font(r, f["body"], f["body_ascii"], size=12)
    for _ in range(4):
        doc.add_paragraph()
    if not args.no_confidential:
        p = doc.add_paragraph()
        r = p.add_run(T["confidential"])
        set_run_font(r, f["head"], f["head_ascii"], size=10, bold=True)
        p = doc.add_paragraph()
        to = (args.to or "").replace("呈送：", "").replace("Prepared for: ", "").strip() or T["to_default"]
        r = p.add_run(T["conf_text"].format(org=args.org or ("本公司" if lang == "zh" else "the Company"), to=to))
        set_run_font(r, f["body"], f["body_ascii"], size=9, color="52514E")
        shade_paragraph(p)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def build_toc(doc, lang, headings):
    """插入 TOC 域；域的缓存结果预填为标题列表（无页码），Word 更新域后会替换为带页码的目录。"""
    f = FONTS[lang]
    T = TEXT[lang]
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(T["toc"])
    set_run_font(r, f["head"], f["head_ascii"], size=16, bold=True)
    entries = [(lvl, txt) for lvl, txt in headings if lvl <= 2] or [(1, "")]
    first = doc.add_paragraph()
    run = first.add_run()
    for tag, extra in (("begin", None), ("instr", 'TOC \\o "1-3" \\h \\z \\u'), ("separate", None)):
        if tag == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = extra
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        run._r.append(el)
    paras = [first]
    for k, (lvl, txt) in enumerate(entries):
        p = first if k == 0 else doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.8 * (lvl - 1))
        r = p.add_run(txt)
        set_run_font(r, f["body"], f["body_ascii"], size=10.5 if lvl == 1 else 10, bold=(lvl == 1))
        paras.append(p)
    last = doc.add_paragraph()
    r = last.add_run(T["update_toc"])
    set_run_font(r, f["body"], f["body_ascii"], size=8.5, color="999999")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    last.add_run()._r.append(end)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def collect_headings(md_text):
    out = []
    for line in md_text.splitlines():
        m = re.match(r"^(#{1,4})\s+(.*)$", line.strip())
        if m:
            out.append((len(m.group(1)), m.group(2).strip()))
    return out


def add_header_footer(doc, title, lang):
    f = FONTS[lang]
    section = doc.sections[0]
    section.different_first_page_header_footer = True  # 封面不显示页眉页脚
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run(title)
    set_run_font(r, f["body"], f["body_ascii"], size=8, color="808080")
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(("第 " if lang == "zh" else "Page "))
    set_run_font(r, f["body"], f["body_ascii"], size=8, color="808080")
    r = add_field(fp, "PAGE")
    set_run_font(r, f["body"], f["body_ascii"], size=8, color="808080")
    r = fp.add_run((" 页" if lang == "zh" else ""))
    set_run_font(r, f["body"], f["body_ascii"], size=8, color="808080")


def parse_table(lines):
    rows = []
    for ln in lines:
        ln = ln.strip()
        if re.match(r"^\|?\s*:?-{2,}", ln.replace("|", "").strip()) or set(ln.replace("|", "").replace(":", "").strip()) <= {"-", " "}:
            continue
        cells = [c.strip() for c in ln.strip("|").split("|")]
        rows.append(cells)
    return rows


def add_table(doc, rows, lang, base_dir):
    f = FONTS[lang]
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_margins(table)
    for i, row in enumerate(rows):
        for j in range(ncols):
            text = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            add_inline(p, text, lang, size=9)
            if i == 0:
                shade_cell(cell)
                for r in p.runs:
                    r.font.bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif NUM_RE.match(text.replace(" ", "")) or text in ("-", "—"):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_caption(doc, text, lang):
    f = FONTS[lang]
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_font(r, f["head"], f["head_ascii"], size=9.5, bold=True, color="333333")


def add_image(doc, caption, path, lang, base_dir):
    full = path if os.path.isabs(path) else os.path.join(base_dir, path)
    if not os.path.exists(full):
        p = doc.add_paragraph()
        add_inline(p, f"[图片缺失：{path}]", lang, color="C00000")
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(full, width=Cm(15))
    if caption:
        add_caption(doc, caption, lang)


def body_paragraph(doc, text, lang, indent=True):
    f = FONTS[lang]
    p = doc.add_paragraph()
    if f["indent"] and indent:
        p.paragraph_format.first_line_indent = Pt(f["body_size"] * 2)
    add_inline(p, text, lang)
    return p


def render_markdown(doc, md_text, lang, base_dir):
    lines = md_text.splitlines()
    i = 0
    para_buf = []
    pending_caption = None

    def flush_para():
        nonlocal para_buf, pending_caption
        if para_buf:
            text = ("" if lang == "zh" else " ").join(s.strip() for s in para_buf)
            if re.match(r"^(表|Table)\s?\d", text):
                pending_caption = text
            else:
                body_paragraph(doc, text, lang)
            para_buf = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped == "" or stripped.startswith("<!--") and "pagebreak" not in stripped:
            flush_para()
            i += 1
            continue
        if "pagebreak" in stripped and stripped.startswith("<!--"):
            flush_para()
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            i += 1
            continue
        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            flush_para()
            pending_caption = None
            level = len(m.group(1))
            doc.add_heading(m.group(2).strip(), level=level)
            i += 1
            continue
        if stripped.startswith("|"):
            flush_para()
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i])
                i += 1
            rows = parse_table(tbl)
            if pending_caption:
                add_caption(doc, pending_caption, lang)
                pending_caption = None
            if rows:
                add_table(doc, rows, lang, base_dir)
            continue
        m = re.match(r"^!\[(.*?)\]\((.*?)\)\s*$", stripped)
        if m:
            flush_para()
            add_image(doc, m.group(1).strip(), m.group(2).strip(), lang, base_dir)
            i += 1
            continue
        m = re.match(r"^(\s*)([-*•]|\d+[.)])\s+(.*)$", line)
        if m:
            flush_para()
            indent = len(m.group(1).replace("\t", "  "))
            ordered = m.group(2)[0].isdigit()
            style = ("List Number" if ordered else "List Bullet") + (" 2" if indent >= 2 else "")
            try:
                p = doc.add_paragraph(style=style)
            except KeyError:
                p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, m.group(3).strip(), lang)
            i += 1
            continue
        if stripped.startswith(">"):
            flush_para()
            quote = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote.append(lines[i].strip().lstrip(">").strip())
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.right_indent = Cm(0.6)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(8)
            add_inline(p, ("" if lang == "zh" else " ").join(q for q in quote if q), lang, size=10)
            shade_paragraph(p, "F4F6FA")
            continue
        if stripped in ("---", "***", "___"):
            flush_para()
            i += 1
            continue
        para_buf.append(line)
        i += 1
    flush_para()


def main():
    ap = argparse.ArgumentParser(description="Markdown → 商业计划书 Word")
    ap.add_argument("markdown")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--title", default=None)
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--org", default="")
    ap.add_argument("--to", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--version", default="")
    ap.add_argument("--lang", default="zh", choices=["zh", "en"])
    ap.add_argument("--logo", default=None)
    ap.add_argument("--no-toc", action="store_true")
    ap.add_argument("--no-cover", action="store_true")
    ap.add_argument("--no-confidential", action="store_true")
    ap.add_argument("--no-chapter-break", action="store_true", help="一级标题不另起一页")
    ap.add_argument("--strip-refs", action="store_true", help="删除正文中的 [A-xx]/[F-xx]/[E-xx]/[S-xx] 标记（对外终稿用；[待补充] 保留）")
    args = ap.parse_args()

    with open(args.markdown, encoding="utf-8") as f:
        md = f.read()
    if args.strip_refs:
        md = re.sub(r"\s?\[(?:[FAES]-?\d+(?:[,，、]\s?[FAES]-?\d+)*|S\d+)\]", "", md)
    base_dir = os.path.dirname(os.path.abspath(args.markdown))
    if not args.title:
        m = re.search(r"^#\s+(.+)$", md, re.M)
        args.title = m.group(1).strip() if m else os.path.splitext(os.path.basename(args.output))[0]

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin = section.bottom_margin = Cm(2.54)
    style_fonts(doc, args.lang)
    add_header_footer(doc, args.title, args.lang)
    if not args.no_cover:
        build_cover(doc, args, args.lang)
    if not args.no_toc:
        build_toc(doc, args.lang, collect_headings(md))
        enable_update_fields(doc)
    if args.no_chapter_break:
        doc.styles["Heading 1"].paragraph_format.page_break_before = False
    render_markdown(doc, md, args.lang, base_dir)
    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    doc.save(args.output)
    print(f"[OK] Word 已生成：{args.output}")
    print("提示：首次用 Word 打开时如提示“是否更新文档中的域”，选择“是”以生成目录；或在目录处右键→更新域。")


if __name__ == "__main__":
    main()
