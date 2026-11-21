# report-ppt-workflow · 汇报 PPT 工作流

![License](https://img.shields.io/github/license/Coder-Lowic/report-ppt-workflow?style=flat-square)
![Skill](https://img.shields.io/badge/Skill-Agent-111111?style=flat-square)
![Claude Code](https://img.shields.io/badge/Claude%20Code-Supported-6B5B95?style=flat-square)
![ZCode](https://img.shields.io/badge/ZCode-Supported-0A7CFF?style=flat-square)
![Codex](https://img.shields.io/badge/Codex-Supported-222222?style=flat-square)
![OpenCode](https://img.shields.io/badge/OpenCode-Supported-1677FF?style=flat-square)

> 🌏 **English: [README.en.md](./README.en.md)**

把一堆杂乱材料变成专业汇报 PPT 的全流程 agent skill:资料整理 → 定制大纲 → 公文写作 → 排版美化。面向体制内、国央企等正式汇报场景,方法论提炼自 [旁门左道PPT](https://www.bilibili.com/video/BV1kLWu64EN3/)。

## 30 秒开始

```bash
npx skills add Coder-Lowic/report-ppt-workflow
```

也可以把这段话直接发给有 shell 权限的 AI:

```text
帮我安装 report-ppt-workflow,运行 npx skills add https://github.com/Coder-Lowic/report-ppt-workflow,装好后告诉我。
```

已安装过,更新用:

```text
帮我更新 report-ppt-workflow,进入 skills 目录对应的 report-ppt-workflow 位置重新安装最新版。
```

装好后直接对 agent 说:

```text
帮我把这份工作总结、台账和几张工作照片做成给领导的季度汇报PPT,10分钟,单位有指定模板(附 pptx)。
```

## 适合 / 不适合

**适合**
- 体制内、国央企等正式工作汇报(政务/年终/季度/专题总结)
- 材料又多又杂,需要先理清逻辑、再定大纲、再写文稿的场景
- 有单位模板/字体约束,要求排版克制严谨、数字有据可查

**不适合**
- 追求海报级视觉创意的 HTML 演示稿(推荐 [guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill))
- 不需要汇报逻辑的单页图/海报

## 工作流

![workflow](assets/workflow.svg)

🎬 **在线交互版**(可缩放/搜索/导出):[coder-lowic.github.io/report-ppt-workflow](https://coder-lowic.github.io/report-ppt-workflow/)


## 依赖(按环境自动升降级)

| 等级 | 条件 | 能力 |
|---|---|---|
| 基础 | 零依赖 | 内容工程全流程 + 排版建议 + 审校清单 |
| 标准 | `pip install python-pptx pillow` | + 模板提取 / 生成可编辑PPT / 讲稿写备注 / 结构验收 |
| 完整 | `pip install pptx-designer` | + 捆绑的 PPT-Design 完整工作流(Build/FreeStyle/VI Build)+ OCR 数字对账 |

PyPI 不可得时,在仓库根目录离线装引擎:`pip install skills/report-ppt-workflow/bundled/vendor/pptx_designer-1.0.0b11-py3-none-any.whl`

## 许可

- 本仓库主体 MIT,见 [LICENSE](LICENSE);第三方组件见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)
- 方法论来源:[旁门左道PPT《领导给一堆材料,怎么用AI又快又好做成PPT?》](https://www.bilibili.com/video/BV1kLWu64EN3/)(BV1kLWu64EN3)
- `bundled/` 内含 [sunchaokun/PPT-Design-Skill](https://github.com/sunchaokun/PPT-Design-Skill) 与 `pptx-designer`(均 MIT)
