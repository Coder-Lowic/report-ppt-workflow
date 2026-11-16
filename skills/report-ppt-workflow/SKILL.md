---
name: report-ppt-workflow
description: 从一堆原始材料(通知/台账/会议纪要/工作总结/数据表/照片说明)产出专业汇报PPT的完整内容工作流:资料整理 → 定制大纲 → 填充文稿 → 排版要点,适配体制内、国央企、职场正式汇报场景。当用户想把材料、文档、数据"做成PPT"、"整理成汇报"、"写汇报大纲"、"做政务/工作/年终/季度总结汇报PPT",或索要"做PPT的提示词"时使用——即使没有明说"工作流"三个字也要触发。Also use when the user asks to "make a report PPT", "turn materials into slides", or wants an outline for a formal work report presentation.
---

# 汇报PPT内容工作流

把一堆杂乱材料变成一份"资料有依据、内容有逻辑、表述严谨、排版克制"的汇报PPT。
提炼自旁门左道PPT《领导给一堆材料,怎么用AI又快又好做成PPT?》(BV1kLWu64EN3,2026-11)。

## 定位与分工

本 skill 负责**内容层**(框架→大纲→文稿→排版建议)。
生成 .pptx 文件的引擎按 `references/design-handoff.md` 第二节选择(首选捆绑 ppt-design 工作流;环境有 `pptx` skill 时亦可走 html2pptx)——先在本文档流程内完成内容,再执行生成。
若用户只要"提示词模板"本身(拿去别的AI用),直接读 `references/prompts.md` 输出模板,不走完整流程。

## 运行环境与能力等级(跨 agent 自持设计)

本 skill 面向任何支持 Agent Skills 规范的 agent(Claude Code/ZCode/Codex/OpenCode/Cursor 等)。核心流程(Step 0–3)是纯知识,零依赖;Step 4 起按环境探测结果选择能力等级,**缺什么降什么,不报错**:

| 等级 | 条件 | 能力 |
|---|---|---|
| **基础** | 零依赖 | Step 0–3 全流程 + 排版建议 + 审校清单(纯文本产出) |
| **标准** | `pip install python-pptx pillow`(任何 agent 环境可装) | + 模板 Token 提取(`scripts/extract_theme.py`)、兜底生成可编辑 PPT(`scripts/minimal_deck.py`)、讲稿写入备注(`scripts/write_notes.py`)、结构验收(`scripts/verify_deck.py` 基础模式) |
| **完整** | `pip install pptx-designer`(MIT,PyPI 公共包);可选 `rapidocr-onnxruntime` | + **自带捆绑的 PPT-Design 完整工作流(`bundled/ppt-design/`,含 Build/FreeStyle/VI Build 三模式与设计规范,能力与上游一致)**、OCR 关键数字对账。**任何 agent 环境,装一个 skill 即得全部能力,零降级** |

> **捆绑组件许可**: `bundled/ppt-design/` 来自 sunchaokun/PPT-Design-Skill(MIT License,见该目录 LICENSE 与 README_UPSTREAM.md),按 MIT 条款捆绑分发并保留其版权声明。其渲染引擎 wheel 已随包分发:`bundled/vendor/pptx_designer-1.0.0b11-py3-none-any.whl`(MIT)——**PyPI 不可得时离线安装:`pip install bundled/vendor/pptx_designer-1.0.0b11-py3-none-any.whl`**。上游仓库/作者服务的存续与本 skill 无关。

> ZCode 生态的 visual-judge(视觉判定)、image-search(搜图)、ppt-design-skill(同引擎的 skill 编排)仅作锦上添花:同会话存在则用,不存在不影响完整级。配图主力始终是用户提供的照片,见 `references/charts-and-images.md`。

探测顺序:先查同会话可用 skill 列表 → 再试 `python -c "import pptx"` → 都没有则按基础级跑并在交付时说明。自带脚本见 `scripts/`(全部为公共 PyPI 依赖或纯标准库,无平台私有工具)。

## 四条核心原则(每一步都适用)

1. **AI干脑力活,人干审美活**:梳理逻辑、找资料缺口、写文稿是AI的事;版式与配色最终由人把关(正式场合模板/字体常被单位锁定)。
2. **数据纪律**:所有数字必须取自用户材料,与原始口径一致;同比注明对比基期;未掌握的数据写"【留白待核】",**严禁估算填充、严禁编造**。这是本工作流的红线。
3. **受众决定材料**:上级领导→成果、价值、战略对齐、资源需求;内部同事→方法、经验、常见问题、协作要点;培训对象→知识、案例、实操。先确认听众,再组织材料。
4. **逻辑主线**:先讲结论 → 再给依据 → 最后提行动。

## 工作流(按序执行,每步结束向用户确认后再进下一步)

### Step 0 · 三问定向
开始前必须确认(材料里能推出来的就不用问):**给谁看?汇报什么?希望达到什么目的?** 顺带收集:汇报场合、时长、页数、模板/字体是否被指定。
材料不全时先列"已有材料清单",再继续——不要求一次凑齐,Step 1 会产出缺料清单。

### Step 0.5 · 模板适配(用户给了单位模板时必做)
用自带脚本 `python scripts/extract_theme.py 模板.pptx` 提取主题色/字体,产出「设计 Token 表」——此后所有展示形式、文稿字号、配色一律从 Token 表取值(细节见 `references/design-handoff.md` 第一节)。没给模板则跳过,用"政务红/深蓝+克制配色"默认建议并在交付时提醒。

### Step 1 · AI整理资料 → 产出「板块框架 + 资料清单」
按 `references/prompts.md` 的【主提示词】结构(角色/任务背景/细节约束/输出格式)组织本次任务的完整上下文,然后产出:
1. **板块框架**:3–5个逻辑板块,含定位说明、时长占比、页数配比(参考配比:成效约55%、问题约15%、请求/行动约25%、封面目录结尾约5%);
2. **资料清单表**:板块标题 | 主要内容(分条) | 需要准备的资料——逐项对照用户已给材料打勾,标出**缺失项**,请用户补充;
3. **咬合自检**:每条问题都要有一条对应的举措或请求"咬合",不能有落空的问题,也不能有无来由的请求。
时间不够时的裁剪原则:删案例页,保住"请求事项"部分的完整度。

### Step 2 · AI定制大纲 → 产出「逐页大纲」
**禁止把材料一股脑塞进一次生成**。在 Step 1 框架被用户确认后,输出逐页大纲,每页只含四项:
`页码 | 页面标题 | 核心内容要点 | 建议展示形式`
展示形式要具体:如"4–6个大数字指标卡(数字+同比箭头)+一句总结导语,避免整段文字"、"对照年度要点逐条标注进度(已完成/推进中)"、"趋势图+数据说明"。
大纲质量标准(反面特征见 `references/examples.md`):不允许出现"在XX领导下坚持XX导向""三个有待""统领—主线—抓手"这类无数据、两页可互换的通稿页。

### Step 3 · AI填充内容 → 产出「每页文稿」
按大纲逐页写正文,规则:
- 每页文字**不超过80字**(口头讲述的细节留给讲稿/备注);
- 数据图表优先于文字段落;图表按三大件型落地(指标卡/趋势图/进度一览表),规范见 `references/charts-and-images.md`;
- **套话改写**:遇到"持续加强、扎实推进、取得良好成效"类表述,按公式 `成果 = 目标 + 动作 + 结果 + 价值` 重写(模板见 references/prompts.md ②);
- **金句**:用户需要时,用"参考文案风格模仿法"生成(模板见 references/prompts.md ③),每条≤30字,参考文案优先用"领导夸过的句子";
- 公文/党建类高严谨场景:提醒用户可换用人民网"写易"引擎(人民日报语料库)复核措辞,本流程产出作底稿;
- 表述对齐用户的政策口径(用户提供什么就 align 什么,没有就不硬造)。

### Step 3.5 · 逐页讲稿 + 计时(口头汇报场景)
按计时公式(成效页40–60秒/页、问题页≤30秒、请求页60–90秒、总时长×0.9)产出每页讲稿,格式与写入 pptx 备注的方法见 `references/review-and-delivery.md` 第二节。仅交书面材料时可跳过。

### Step 4 · 排版、验收与交付
**满页文字页**(内容不能删时):万能四步分层法——**分组**(按主题归组)→**提炼**(小标题)→**突出**(放大关键数据)→**留白**(拉开层次),完整案例见 references/examples.md;
**配色**:HSL 十字正交法——取一个主题色,向上调高亮度找浅色、向下调低亮度找深色(纵向成深浅层次),横向微调色调(H)取相邻色(需要多组区分时用),整体克制不花哨;有 Token 表时只在表内取值。
**生成文件**:需要 .pptx 时按 `references/design-handoff.md` 第二节引擎选择——**首选:读 `bundled/ppt-design/SKILL.md` 并按其完整工作流执行(brief→视觉方向→ThemeLock→Build→渲染→QA,设计能力最强)**;快速稿用 `scripts/design_deck.py`;兜底用 `scripts/minimal_deck.py`;
**视觉验收**:文件生成后必做——`python scripts/verify_deck.py deck.pptx`(越界/空页/字量检查,任何环境)+ 渲染 PNG(`scripts/export_deck.ps1`/`.sh`)后 `--ocr png目录 --facts 关键数字` 对账(需 rapidocr);agent 有看图能力时逐页复查,fail 页修复复验;
**审校放行**:按 `references/review-and-delivery.md` 的 10 项清单逐项过,数据类任何一项不过=打回,不出门;
**交付物**(问用户选型):① Markdown 大纲+文稿;② .pptx(+PDF/讲义版/图片版,导出命令见 review-and-delivery.md);③ 只要提示词模板(references/prompts.md)。

## 何时读 references

| 场景 | 读哪个 |
|---|---|
| 组装 Step 1 提示词上下文 / 用户要提示词模板 | `references/prompts.md` |
| 拿不准大纲或文稿的好坏标准 / 需要正反示例 | `references/examples.md` |
| 用户给了单位模板 / 要生成 .pptx / 需要视觉验收 | `references/design-handoff.md` |
| 台账数据要变图表 / 需要配图、照片墙排版 | `references/charts-and-images.md` |
| 交付前审校 / 写讲稿与计时 / 导出 PDF 讲义版 | `references/review-and-delivery.md` |
| 需要执行模板提取/兜底生成/写备注/导出 | `scripts/`(extract_theme.py / minimal_deck.py / write_notes.py / export_deck.ps1/.sh) |
