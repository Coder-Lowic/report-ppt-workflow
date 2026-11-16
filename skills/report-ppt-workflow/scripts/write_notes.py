# -*- coding: utf-8 -*-
"""把逐页讲稿写入 pptx 演讲者备注。依赖:pip install python-pptx

用法: python write_notes.py deck.pptx notes.json
notes.json 格式(页码从1开始): {"1": "[45秒] 开口第一句...", "2": "..."}
"""
import json
import sys

from pptx import Presentation


def main(pptx_path, notes_path):
    notes = json.loads(open(notes_path, encoding="utf-8").read())
    prs = Presentation(pptx_path)
    hits, misses = 0, []
    for i, slide in enumerate(prs.slides, start=1):
        text = notes.get(str(i))
        if not text:
            misses.append(i)
            continue
        slide.notes_slide.notes_text_frame.text = text
        hits += 1
    prs.save(pptx_path)
    print(f"已写入 {hits} 页备注;未提供讲稿的页:{misses or '无'}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
