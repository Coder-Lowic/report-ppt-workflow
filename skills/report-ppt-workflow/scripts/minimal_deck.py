# -*- coding: utf-8 -*-
"""兜底 PPT 生成器:无任何设计 skill 时,纯 python-pptx 产出结构化、完全可编辑的汇报 PPT。
依赖: pip install python-pptx

用法: python minimal_deck.py outline.json -o deck.pptx

outline.json 结构:
{
  "title": "2026年前三季度工作汇报", "subtitle": "某单位 · 专题推进会",
  "primary": "#C00000",            # 可选,默认政务红;单位有模板请先用 extract_theme.py 提取
  "slides": [
    {"type": "cover", "title": "...", "subtitle": "...", "footer": "汇报单位 2026-11"},
    {"type": "toc",   "title": "目录", "items": ["总体情况", "重点成效", ...]},
    {"type": "kpi",   "title": "核心指标(2026年1-9月)",
     "cards": [{"value": "258,050", "label": "受理总量"}, {"value": "99.46%", "label": "按时办结率"}],
     "note": "数据说明:取自月度台账"},
    {"type": "bullets", "title": "重点成效(一)", "items": ["…≤80字/页…"]},
    {"type": "table",  "title": "年度重点工作推进一览",
     "headers": ["事项", "年度目标", "当前进度", "状态"],
     "rows": [["…", "…", "…", "已完成"]]}
  ]
}
"""
import json
import re
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

EMU_W, EMU_H = Inches(13.333), Inches(7.5)
FONT = "微软雅黑"
GOV_RED = RGBColor(0xC0, 0x00, 0x00)
MUTED = RGBColor(0x99, 0x99, 0x99)
# 句内数字片段(含单位),用于要点页关键数字强调
NUM_RE = re.compile(r"\d[\d,\.]*\s*(?:%|项|件|次|分钟|小时|个|家|条|页|万|亿|元)?")
STATUS_STYLE = {  # 表格状态列语义着色(白字)
    "已完成": None,   # 动态填 primary
    "推进中": RGBColor(0x59, 0x59, 0x59),
    "滞后": GOV_RED,
}


def hexrgb(h, default="C00000"):
    h = (h or default).lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_text(slide, x, y, w, h, text, size, bold=False, color=None, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in str(text).split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        r = p.add_run()
        r.text = line
        f = r.font
        f.name, f.size, f.bold = FONT, Pt(size), bold
        f.color.rgb = color or RGBColor(0x26, 0x26, 0x26)
    return box


def add_rich(slide, x, y, w, h, text, size, primary):
    """要点文本:句内数字片段自动加粗、主题色、放大2pt(政务'关键数字醒目'惯例)。"""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    pos = 0
    for m in NUM_RE.finditer(text):
        if m.start() > pos:
            r = p.add_run()
            r.text = text[pos:m.start()]
            r.font.name, r.font.size = FONT, Pt(size)
            r.font.color.rgb = RGBColor(0x26, 0x26, 0x26)
        r = p.add_run()
        r.text = m.group()
        r.font.name, r.font.size, r.font.bold = FONT, Pt(size + 2), True
        r.font.color.rgb = primary
        pos = m.end()
    if pos < len(text):
        r = p.add_run()
        r.text = text[pos:]
        r.font.name, r.font.size = FONT, Pt(size)
        r.font.color.rgb = RGBColor(0x26, 0x26, 0x26)


def header(slide, primary, title):
    bar = slide.shapes.add_shape(1, 0, 0, EMU_W, Inches(0.9))  # 顶栏
    bar.fill.solid()
    bar.fill.fore_color.rgb = primary
    bar.line.fill.background()
    add_text(slide, Inches(0.5), Inches(0.15), EMU_W - Inches(1), Inches(0.6),
             title, 28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))


def page_no(slide, idx, total):
    """页码 N/M,右下角,封面免(政务惯例)。"""
    add_text(slide, EMU_W - Inches(1.2), EMU_H - Inches(0.45), Inches(1), Inches(0.3),
             f"{idx}/{total}", 10, color=MUTED, align=PP_ALIGN.RIGHT)


def render(prs, deck):
    primary = hexrgb(deck.get("primary"))
    slides = deck["slides"]
    total = len(slides)
    for s_idx, s in enumerate(slides, 1):
        t = s["type"]
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
        if t == "cover":
            band = slide.shapes.add_shape(1, 0, EMU_H - Inches(1.1), EMU_W, Inches(1.1))
            band.fill.solid()
            band.fill.fore_color.rgb = primary
            band.line.fill.background()
            title_text = s.get("title", deck.get("title", ""))
            tsize = 40 if len(title_text) <= 16 else (36 if len(title_text) <= 24 else 30)
            add_text(slide, Inches(0.8), Inches(2.2), EMU_W - Inches(1.6), Inches(1.2),
                     title_text, tsize, bold=True, color=primary)
            add_text(slide, Inches(0.8), Inches(3.5), EMU_W - Inches(1.6), Inches(0.8),
                     s.get("subtitle", deck.get("subtitle", "")), 20)
            if s.get("footer"):
                add_text(slide, Inches(0.8), EMU_H - Inches(0.85), EMU_W - Inches(1.6),
                         Inches(0.5), s["footer"], 14, color=RGBColor(0xFF, 0xFF, 0xFF))
        elif t == "toc":
            header(slide, primary, s.get("title", "目录"))
            for i, item in enumerate(s["items"], 1):
                num = slide.shapes.add_shape(5, Inches(0.8), Inches(1.5 + i * 0.95),
                                             Inches(0.5), Inches(0.5))  # 圆
                num.fill.solid()
                num.fill.fore_color.rgb = primary
                num.line.fill.background()
                add_text(slide, Inches(0.8), Inches(1.53 + i * 0.95), Inches(0.5), Inches(0.45),
                         str(i).zfill(2), 16, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
                         align=PP_ALIGN.CENTER)
                add_text(slide, Inches(1.5), Inches(1.5 + i * 0.95), Inches(10), Inches(0.5),
                         item, 20)
        elif t == "kpi":
            header(slide, primary, s["title"])
            cards = s["cards"][:6]
            cols = 3 if len(cards) > 3 else len(cards)
            rows = (len(cards) + cols - 1) // cols
            gap, cw, ch = Inches(0.3), (EMU_W - Inches(1) - Inches(0.3) * (cols - 1)) / cols, Inches(2.1)
            for i, c in enumerate(cards):
                cx = Inches(0.5) + (i % cols) * (cw + gap)
                cy = Inches(1.5) + (i // cols) * (ch + gap)
                card = slide.shapes.add_shape(5, cx, cy, cw, ch)
                card.fill.solid()
                card.fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
                card.line.color.rgb = primary
                add_text(slide, cx, cy + Inches(0.25), cw, Inches(0.9), c["value"], 34,
                         bold=True, color=primary, align=PP_ALIGN.CENTER)
                add_text(slide, cx, cy + Inches(1.3), cw, Inches(0.5), c["label"], 15,
                         align=PP_ALIGN.CENTER)
            if s.get("note"):
                add_text(slide, Inches(0.5), EMU_H - Inches(0.6), EMU_W - Inches(1),
                         Inches(0.4), s["note"], 11, color=RGBColor(0x80, 0x80, 0x80))
        elif t == "bullets":
            header(slide, primary, s["title"])
            for i, item in enumerate(s["items"][:8], 1):
                add_text(slide, Inches(0.7), Inches(1.4 + i * 0.72), Inches(0.4), Inches(0.5),
                         "·", 22, bold=True, color=primary)
                add_rich(slide, Inches(1.15), Inches(1.4 + i * 0.72), Inches(11.5),
                         Inches(0.65), item, 17, primary)
        elif t == "table":
            header(slide, primary, s["title"])
            rows, cols = len(s["rows"]) + 1, len(s["headers"])
            tbl = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.4),
                                         EMU_W - Inches(1), Inches(0.5 * rows)).table
            for j, htext in enumerate(s["headers"]):
                cell = tbl.cell(0, j)
                cell.text = htext
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.bold, r.font.size, r.font.name = True, Pt(15), FONT
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cell.fill.solid()
                cell.fill.fore_color.rgb = primary
            for i, row in enumerate(s["rows"], 1):
                for j, val in enumerate(row):
                    cell = tbl.cell(i, j)
                    cell.text = str(val)
                    status_bg = STATUS_STYLE.get(str(val).strip())
                    if status_bg is None and str(val).strip() == "已完成":
                        status_bg = primary
                    if status_bg is not None:  # 状态单元格语义着色+白字(徽章惯例)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = status_bg
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.size, r.font.name = Pt(14), FONT
                            if status_bg is not None:
                                r.font.bold = True
                                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if t != "cover":
            page_no(slide, s_idx, total)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    out = sys.argv[sys.argv.index("-o") + 1] if "-o" in sys.argv else "deck.pptx"
    deck = json.loads(open(args[0], encoding="utf-8").read())
    prs = Presentation()
    prs.slide_width, prs.slide_height = EMU_W, EMU_H
    render(prs, deck)
    prs.save(out)
    print(f"已生成 {out}({len(deck['slides'])} 页,全部原生可编辑元素)")


if __name__ == "__main__":
    main()
