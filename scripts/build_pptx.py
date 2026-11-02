#!/usr/bin/env python3
"""
build_pptx.py — 把 deck.md 做成简洁的路演 PPT（环境里有更强的 pptx skill 时优先用那个，本脚本是通用备选）。

用法:
    python build_pptx.py deck.md -o "项目-路演PPT-2026-11.pptx" [--lang zh|en] [--accent 2F5496]

deck.md 格式（每页以 # 开头）:
    # 封面标题            ← 第一页为封面；下一行 "## 副标题" 可选；"组织/日期" 写成普通行
    # 页面标题
    要点一                 ← 普通行或 "- 要点" 都作为要点
    - 要点二
    ![](charts/revenue_profit.png)   ← 图片放右侧（有图时要点在左），按比例缩放到可用区域
    | 指标 | 2027 | 2028 |          ← Markdown 表格会转成 PPT 表格（要点在上、表在下）
    > 讲稿备注              ← 进入演讲者备注
    字号随要点数量自动调整；每页建议 3–5 条要点、每条一行，超过 8 条脚本会提示拆页。
    ---                      ← 可选的分隔，忽略
"""
import argparse
import os
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt

FONT = {"zh": "微软雅黑", "en": "Calibri"}


def parse_deck(md):
    slides = []
    cur = None
    for line in md.splitlines():
        s = line.rstrip()
        if s.startswith("# "):
            cur = {"title": s[2:].strip(), "subtitle": None, "bullets": [], "image": None, "notes": []}
            slides.append(cur)
            continue
        if cur is None:
            continue
        if s.startswith("## "):
            cur["subtitle"] = s[3:].strip()
        elif s.strip().startswith(">"):
            cur["notes"].append(s.strip().lstrip(">").strip())
        else:
            m = re.match(r"^!\[(.*?)\]\((.*?)\)", s.strip())
            if m:
                cur["image"] = m.group(2).strip()
            elif s.strip().startswith("|"):
                cells = [c.strip() for c in s.strip().strip("|").split("|")]
                if set("".join(cells)) <= set("-: "):
                    continue
                cur.setdefault("table", []).append(cells)
            elif s.strip() in ("", "---"):
                continue
            else:
                indent = len(s) - len(s.lstrip())
                txt = re.sub(r"^[-*•]\s+|^\d+[.)]\s+", "", s.strip())
                cur["bullets"].append((1 if indent >= 2 else 0, txt))
    return slides


def fit_sizes(bullets):
    """按文字量选字号：要点少字大，要点多字小；超过 8 条或 400 字提示精简。"""
    total = sum(len(b) for _, b in bullets)
    if len(bullets) <= 4 and total <= 160:
        return 20, 16
    if len(bullets) <= 6 and total <= 300:
        return 17, 14
    if len(bullets) <= 8 and total <= 420:
        return 15, 13
    return 13, 11


def add_text(frame, text, size, bold=False, color=None, font="微软雅黑", align=None, level=0):
    frame.word_wrap = True
    p = frame.paragraphs[0] if (len(frame.paragraphs) == 1 and not frame.paragraphs[0].text) else frame.add_paragraph()
    p.level = level
    p.alignment = align or PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = font
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    p.space_after = Pt(8)
    return p


def build(slides, out, lang, accent, base_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)  # 16:9
    blank = prs.slide_layouts[6]
    font = FONT[lang]
    W, H = prs.slide_width, prs.slide_height
    for idx, sd in enumerate(slides):
        slide = prs.slides.add_slide(blank)
        if idx == 0:
            bg = slide.shapes.add_shape(1, 0, 0, W, H)
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor.from_string(accent)
            bg.line.fill.background()
            tb = slide.shapes.add_textbox(Cm(2.5), Cm(6), W - Cm(5), Cm(4))
            add_text(tb.text_frame, sd["title"], 40, bold=True, color="FFFFFF", font=font)
            if sd["subtitle"]:
                add_text(tb.text_frame, sd["subtitle"], 20, color="E8EEF8", font=font)
            if sd["bullets"]:
                tb2 = slide.shapes.add_textbox(Cm(2.5), Cm(13), W - Cm(5), Cm(4))
                for _, b in sd["bullets"]:
                    add_text(tb2.text_frame, b, 14, color="E8EEF8", font=font)
        else:
            bar = slide.shapes.add_shape(1, 0, 0, Cm(0.6), H)
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor.from_string(accent)
            bar.line.fill.background()
            tb = slide.shapes.add_textbox(Cm(2), Cm(1.0), W - Cm(4), Cm(2.6))
            title_size = 28 if len(sd["title"]) <= 22 else (24 if len(sd["title"]) <= 30 else 20)
            add_text(tb.text_frame, sd["title"], title_size, bold=True, color="1A1A1A", font=font)
            has_img = sd["image"] and os.path.exists(sd["image"] if os.path.isabs(sd["image"]) else os.path.join(base_dir, sd["image"]))
            has_tbl = bool(sd.get("table"))
            body_w = (W - Cm(4)) * (0.46 if has_img else 1)
            body_h = (H - Cm(6)) * (0.5 if (has_tbl and sd["bullets"]) else 1)
            size0, size1 = fit_sizes(sd["bullets"])
            if len(sd["bullets"]) > 8 or sum(len(b) for _, b in sd["bullets"]) > 420:
                print(f"[提示] 第 {idx + 1} 页要点过多（{len(sd['bullets'])} 条），建议拆页或精简")
            if sd["bullets"] or sd["subtitle"]:
                body = slide.shapes.add_textbox(Cm(2), Cm(4), body_w, body_h)
                tf = body.text_frame
                tf.word_wrap = True
                if sd["subtitle"]:
                    add_text(tf, sd["subtitle"], 15, color="52514E", font=font)
                for level, b in sd["bullets"]:
                    add_text(tf, ("• " if level == 0 else "– ") + b, size0 if level == 0 else size1, color="1A1A1A", font=font, level=level)
            if has_tbl:
                rows = sd["table"]
                ncols = max(len(r) for r in rows)
                top = Cm(4) + (body_h + Cm(0.4) if sd["bullets"] else 0)
                tbl_h = min(H - top - Cm(1.5), Cm(0.9) * len(rows))
                shape = slide.shapes.add_table(len(rows), ncols, Cm(2), top, body_w if has_img else W - Cm(4), tbl_h)
                tbl = shape.table
                for ri, row in enumerate(rows):
                    for ci in range(ncols):
                        cell = tbl.cell(ri, ci)
                        cell.text = row[ci] if ci < len(row) else ""
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                r.font.size = Pt(12 if len(rows) <= 6 else 10)
                                r.font.name = font
                                r.font.bold = (ri == 0)
                            if ci > 0:
                                p.alignment = PP_ALIGN.RIGHT
            if has_img:
                path = sd["image"] if os.path.isabs(sd["image"]) else os.path.join(base_dir, sd["image"])
                from PIL import Image as _Img
                iw, ih = _Img.open(path).size
                max_w = W - Cm(4) - body_w - Cm(0.8)
                max_h = H - Cm(5.5)
                scale = min(max_w / iw, max_h / ih)
                pic_w, pic_h = int(iw * scale), int(ih * scale)
                slide.shapes.add_picture(path, Cm(2) + body_w + Cm(0.8) + (max_w - pic_w) // 2, Cm(4.2), width=pic_w, height=pic_h)
            num = slide.shapes.add_textbox(W - Cm(3), H - Cm(1.4), Cm(2), Cm(1))
            add_text(num.text_frame, str(idx + 1), 10, color="999999", font=font, align=PP_ALIGN.RIGHT)
        if sd["notes"]:
            slide.notes_slide.notes_text_frame.text = "\n".join(sd["notes"])
    prs.save(out)
    print(f"[OK] PPT 已生成：{out}（{len(slides)} 页）")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--lang", default="zh", choices=["zh", "en"])
    ap.add_argument("--accent", default="2F5496")
    args = ap.parse_args()
    with open(args.deck, encoding="utf-8") as f:
        slides = parse_deck(f.read())
    if not slides:
        raise SystemExit("deck.md 中没有以 # 开头的页面")
    build(slides, args.output, args.lang, args.accent, os.path.dirname(os.path.abspath(args.deck)))


if __name__ == "__main__":
    main()
