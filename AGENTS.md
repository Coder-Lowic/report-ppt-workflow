# AGENTS.md

面向在本仓库工作的 AI 编码 agent 的说明。

## 仓库布局

- `skills/report-ppt-workflow/` — skill 本体,所有修改都在这里:
  - `SKILL.md` 工作流主文档(触发逻辑、能力分级、步骤编排)
  - `references/` 按需加载的规范文档(提示词模板/正反示例/设计交接/图表规范/审校交付)
  - `scripts/` 自带工具链(Python/PowerShell/Bash,公共 PyPI 依赖或纯标准库)
  - `bundled/ppt-design/` 第三方 PPT-Design-Skill 副本(MIT)——**不要修改**,更新见下
  - `bundled/vendor/` pptx-designer wheel(MIT)——**不要修改**
- `assets/` 工作流配图(SVG)
- 根目录:README / README.en / LICENSE / CHANGELOG / THIRD_PARTY_NOTICES

## 修改守则

- skill 的行为规范改动写进 `SKILL.md` 或对应 reference,保持"主文档 <500 行、细节放 references"的渐进披露结构
- 脚本改动后必须实测(生成→渲染→验收全链路至少跑一遍)
- 更新捆绑组件:从上游 `git clone --depth 1` 后替换 `bundled/ppt-design/` 并保留其 LICENSE,同步更新 THIRD_PARTY_NOTICES.md 与 CHANGELOG.md
- 提交信息用一句话说清改了什么(中文或英文均可)

## 本地验证

```bash
npx skills add ./skills/report-ppt-workflow   # 本地安装到当前 agent 验证触发
```
