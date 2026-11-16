# -*- coding: utf-8 -*-
"""视觉验收(开源版):不依赖任何 agent 生态的确定性检查,替代 visual-judge 的核心功能。

依赖: python-pptx(必需);rapidocr_onnxruntime(可选,启用 --ocr 数字对账)
用法:
  python verify_deck.py deck.pptx                       # 结构检查:越界/空页/每页字数
  python verify_deck.py deck.pptx --ocr png目录 --facts "258,050" "99.46%"   # +数字对账

退出码: 0=全部通过, 1=有 fail 项(修复后复验)
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu


def check_structure(path):
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    issues, warns = [], []
    for i, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        if not shapes:
            issues.append(f"第{i}页:空白页(无任何元素)")
            continue
        chars = 0
        for sh in shapes:
            try:  # 越界检测:含文字元素越界=fail;无文字形状(出血装饰)仅warn
                if sh.left is None:
                    continue
                oob = sh.left < -Emu(91440) or sh.top < -Emu(91440) \
                    or sh.left + sh.width > W + Emu(91440) or sh.top + sh.height > H + Emu(91440)
                if oob:
                    has_text = sh.has_text_frame and sh.text_frame.text.strip()
                    msg = (f"第{i}页:元素「{sh.shape_type}」越出画布 "
                           f"(left={Emu(sh.left).inches:.1f}in top={Emu(sh.top).inches:.1f}in)")
                    (issues if has_text else warns).append(msg)
            except Exception:
                pass
            if sh.has_text_frame:
                chars += sum(len(p.text) for p in sh.text_frame.paragraphs)
        if chars > 80:  # 政务规范:每页正文不超80字(图表页表格另计,先warn)
            warns.append(f"第{i}页:文字量 {chars} 字(>80,建议四步分层法改造)")
        print(f"  第{i}页: {len(shapes)} 元素, 文字 {chars} 字")
    return issues, warns


def check_ocr(png_dir, facts):
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        print("  [跳过] 未装 rapidocr_onnxruntime(pip install rapidocr-onnxruntime 启用)")
        return []
    ocr = RapidOCR()

    def _text_score(r):
        # 兼容 rapidocr 新旧返回结构:[box, text, score] 或 {'text','score'} dict
        if isinstance(r, dict):
            return r.get("text", ""), float(r.get("score", 1.0))
        return r[1], float(r[2]) if len(r) > 2 else 1.0

    text_all = []
    for png in sorted(Path(png_dir).glob("*.png")) + sorted(Path(png_dir).glob("*.PNG")):
        result, _ = ocr(str(png))
        text_all += [t for t, s in (_text_score(r) for r in (result or [])) if s > 0.5]
    # 归一化:OCR 常把千分位逗号识别为点,故逗号/点/空格一并去除后比对
    norm = lambda s: "".join(c for c in str(s) if c not in " ,.")
    joined = norm("".join(text_all))
    missing = [f for f in facts if norm(f) not in joined]
    return [f"OCR 对账:关键数字「{m}」未在渲染页面中出现(核对是否漏页/被截断/写错)" for m in missing]


def main():
    path = sys.argv[1]
    print(f"结构检查:{path}")
    issues, warns = check_structure(path)
    if "--ocr" in sys.argv:
        png_dir = sys.argv[sys.argv.index("--ocr") + 1]
        rest = sys.argv[sys.argv.index("--facts") + 2:] if "--facts" in sys.argv else []
        facts = [f for f in rest if not f.startswith("-")]
        print(f"OCR 数字对账({len(facts)} 个关键数字):")
        issues += check_ocr(png_dir, facts)
    print()
    for w in warns:
        print("WARN", w)
    for e in issues:
        print("FAIL", e)
    verdict = "通过" if not issues else f"未通过({len(issues)} 项 fail)"
    print(f"\n结论:{verdict}" + (f",{len(warns)} 项提醒" if warns else ""))
    sys.exit(0 if not issues else 1)


if __name__ == "__main__":
    main()
