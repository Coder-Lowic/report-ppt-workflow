# -*- coding: utf-8 -*-
"""从单位模板 pptx 提取设计 Token(主题色/字体),零第三方依赖(仅标准库)。

用法: python extract_theme.py 模板.pptx
输出: 设计 Token 表(markdown),供 report-ppt-workflow 全流程引用
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ROLES = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
         "accent4", "accent5", "accent6", "hlink", "folHlink"]
ROLE_CN = {"dk1": "深色1(正文)", "lt1": "浅色1(背景)", "dk2": "深色2",
           "lt2": "浅色2", "hlink": "超链接", "folHlink": "已访问链接"}


def clr_of(el):
    """读取 <a:srgbClr>/<a:sysClr> 颜色值。"""
    for child in el:
        tag = child.tag.split("}")[1]
        if tag == "srgbClr":
            return "#" + child.get("val", "")
        if tag == "sysClr":
            return "#" + child.get("lastClr", "000000")
    return ""


def main(path):
    with zipfile.ZipFile(path) as z:
        theme_xml = z.read("ppt/theme/theme1.xml")
    root = ET.fromstring(theme_xml)

    scheme = root.find(f".//{{{A}}}clrScheme")
    name = scheme.get("name", "") if scheme is not None else ""
    print(f"# 设计 Token 表(来源:{path} 主题:{name})\n")
    print("## 主题色\n| 角色 | 色值 | 建议用途 |")
    print("|---|---|---|")
    usage = {"accent1": "主色:标题条/强调/图表主系列",
             "accent2": "辅色:第二数据系列/次级强调",
             "accent3": "辅色2", "accent4": "备用", "accent5": "备用",
             "accent6": "备用", "dk1": "正文字色", "lt1": "页面背景",
             "dk2": "标题深色变体", "lt2": "卡片/分区底色"}
    for role in ROLES:
        el = scheme.find(f"{{{A}}}{role}") if scheme is not None else None
        val = clr_of(el) if el is not None else ""
        label = ROLE_CN.get(role, role)
        print(f"| {label} | {val} | {usage.get(role, '')} |")

    fonts = root.find(f".//{{{A}}}fontScheme")
    print("\n## 字体规范")
    if fonts is None:
        print("(未找到 fontScheme)")
        return
    for part, cn in (("majorFont", "标题字体"), ("minorFont", "正文字体")):
        f = fonts.find(f"{{{A}}}{part}")
        if f is None:
            continue
        latin = f.find(f"{{{A}}}latin")
        ea = f.find(f"{{{A}}}ea")
        print(f"- **{cn}**:西文 `{latin.get('typeface') if latin is not None else '?'}`"
              f" / 中文 `{ea.get('typeface') if ea is not None else '?'}`")
    print("\n> 字号规范请打开模板实测(或让 agent 解包看 slides/slide1.xml 的 <a:rPr sz>)。")
    print("> 后续所有展示形式、文稿字号、配色一律从本表取值,不引入表外颜色。")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "模板.pptx")
