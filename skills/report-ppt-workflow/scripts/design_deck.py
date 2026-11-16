# -*- coding: utf-8 -*-
"""设计引擎(开源版):用 PyPI 的 pptx-designer(MIT)生成高质感可编辑 PPT。
这是 ppt-design-skill 的开源替代——只依赖公共包,任何 agent 环境可用。

依赖: pip install pptx-designer
用法:
  python design_deck.py outline.json -o deck.pptx                          # Build模式(默认,可控)
  python design_deck.py outline.json --template 单位模板.pptx -o deck.pptx # VI 模板合规
  python design_deck.py outline.json --mode freestyle                      # FreeStyle(引擎自动规划)
  python design_deck.py outline.json --palette government-public-service --seed 42
outline.json 即 minimal_deck.py 的输入格式;页型映射:cover→hook, toc→hook,
kpi→data, bullets→features/solution, table→data(引擎无原生表格,复杂数据表用 minimal_deck)。
配合 export_deck + verify_deck 使用:生成→渲染→审查→调参重生成(见 design-handoff.md)。
"""
import argparse
import json

from pptx_designer import PALETTES, STYLES, generate_ppt

GOAL_MAP = {
    "cover": "hook",        # 封面(只渲染标题/副标题,不渲染bullets)
    "toc": "features",      # 目录(features版式逐条渲染列表)
    "kpi": "data",          # 指标页(METRICS版式,渲染"标签:数值"条目)
    "bullets": "features",  # 要点列举
    "table": "features",    # 引擎无原生表格;转为逐行要点(规范表格建议走 minimal_deck)
}


def outline_to_content(deck):
    """outline.json → generate_ppt 的 content 结构(title + pages[goal/title/subtitle/bullets])。"""
    pages = []
    for s in deck.get("slides", []):
        t = s.get("type", "bullets")
        page = {"goal": GOAL_MAP.get(t, "features"), "title": s.get("title", "")}
        if t == "cover":
            sub = s.get("subtitle", deck.get("subtitle", ""))
            if s.get("footer"):  # hook 版式丢弃 bullets,落款并入副标题
                sub = f"{sub} · {s['footer']}" if sub else s["footer"]
            page.update(title=s.get("title", deck.get("title", "")), subtitle=sub, bullets=[])
        elif t == "toc":
            page.update(subtitle="汇报提纲", bullets=[f"{i}. {x}" for i, x in enumerate(s.get("items", []), 1)])
        elif t == "kpi":
            # 引擎硬约束:METRICS 版式>4条整页静默空白;分隔必须为": "(冒号+空格)才拆成值/标签
            cards = [f"{c['label']}: {c['value']}" for c in s.get("cards", [])][:4]
            page.update(subtitle=(s.get("note") or "核心指标一览")[:24], bullets=cards)
        elif t == "bullets":
            page.update(subtitle="", bullets=s.get("items", [])[:6])
        elif t == "table":
            hdr = s.get("headers", [])
            rows = [", ".join(f"{h}:{v}" for h, v in zip(hdr, r) if str(v).strip()) for r in s.get("rows", [])]
            page.update(subtitle="年度重点工作进度", bullets=rows[:6])
        pages.append(page)
    return {"title": deck.get("title", ""), "pages": pages}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("outline")
    ap.add_argument("-o", "--output", default="deck.pptx")
    ap.add_argument("--mode", choices=["build", "freestyle"], default="build")
    ap.add_argument("--style", default="minimalism-swiss-style",
                    help=f"设计风格,84种,如 {', '.join(list(STYLES)[:5])}…")
    ap.add_argument("--palette", default="midnight-navy",
                    help="配色(仅 COLOR_PALETTES 内30个有效,其他名会静默回落 ocean-blue);"
                         "政务推荐:midnight-navy/cherry-red/wine-burgundy/ocean-blue/ink-wash")
    ap.add_argument("--template", default=None, help="单位模板.pptx(VI 模板合规)")
    ap.add_argument("--seed", type=int, default=None, help="风格种子(复现/换一版)")
    a = ap.parse_args()

    deck = json.loads(open(a.outline, encoding="utf-8").read())
    kwargs = dict(style=a.style, palette=a.palette, output=a.output, seed=a.seed)
    if a.template:
        kwargs["template"] = a.template
    if a.mode == "build":
        kwargs["content"] = outline_to_content(deck)
    else:
        kwargs["query"] = "\n".join([deck.get("title", "")] + [s.get("title", "") for s in deck.get("slides", [])])
    r = generate_ppt(**kwargs)
    print(f"已生成 {r.get('output_path', a.output)}:{r.get('slide_count', '?')} 页 / "
          f"{r.get('shapes_count', '?')} 个元素 / 主题 {r.get('theme', '')}(全部原生可编辑)")
    if r.get("theme_atoms"):
        print("组合:", r["theme_atoms"])


if __name__ == "__main__":
    main()
