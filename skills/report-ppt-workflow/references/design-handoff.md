# 设计交接(report-ppt-workflow)

内容层(本 skill)→ 视觉层的交接规范:模板适配、设计引擎调用、视觉验收。

---

## 一、模板适配(Step 0.5,用户提供了单位模板时必做)

正式场合(体制内/国央企)模板、字体常被锁定。**在产出任何大纲建议展示形式之前**,先提取模板的设计要素并全程对齐。

### 提取方法(自带脚本优先,零平台依赖)

```bash
# 首选:自带脚本(纯标准库,直接输出 Token 表)
python scripts/extract_theme.py 模板.pptx
```

需要更深入信息(母版用色习惯、实际字号)时再手动解包:

```bash
# 解包模板文件(zip 结构,可用任何 unzip 工具)
python -c "import zipfile; zipfile.ZipFile('模板.pptx').extractall('模板解包')"

# ppt/theme/theme1.xml → 主题色 <a:clrScheme>、字体 <a:fontScheme>(脚本已解析)
# ppt/slideMasters/slideMaster1.xml → grep <a:srgbClr> 统计实际用色
# ppt/slides/slide1.xml → 字体字号 <a:rPr sz="...">
```

### 产出「设计 Token 表」(后续所有步骤引用)

| Token | 来源 | 示例 |
|---|---|---|
| 主色 | clrScheme accent1 或母版高频色 | #C00000(政务红) |
| 辅色 | accent2-3 / 深浅变化 | 深红#8B0000、浅红#F2D5D5 |
| 标题字体 | fontScheme major | 方正小标宋 / 黑体 |
| 正文字体 | fontScheme minor | 仿宋_GB2312 / 微软雅黑 |
| 字号规范 | slide1.xml 实测 | 标题32pt、一级20pt、正文16pt |
| 版式禁区 | 观察母版 | 页眉Logo不可遮挡、页脚页码位置 |

规则:
- Step 2 大纲的"建议展示形式"、Step 3 的每页文稿字号、Step 4 配色,**一律从 Token 表取值**,不引入模板外颜色;
- 模板锁死时,本 skill 的 HSL 十字正交法只在主色的深浅/相邻范围内使用;
- 用户没给模板 → 跳过本步,用"政务红/深蓝 + 克制配色"默认建议,并在交付时提醒"如单位有指定模板请提供,可一键对齐"。

---

## 二、设计引擎调用(Step 4)——按需求选引擎

### ① 首选:捆绑的 PPT-Design 完整工作流(设计能力最强,自带零依赖装)
**读 `bundled/ppt-design/SKILL.md`,按其工作流执行**:brief 契约 → 视觉方向三选一 → Theme Lock 钉定 → Build Mode 生成(其 references/ 有设计规范与 QA 文档,scripts/ 有渲染与校验工具)→ PPTX→PDF→PNG → 审查修订。
- 能力与独立安装上游 skill 完全一致(Build/FreeStyle/VI Build 三模式);
- 需要 `pip install pptx-designer`(引擎运行时);
- 本机若同会话装有独立 `ppt-design-skill`,两者同源,任用其一;
- 对照实验结论(2026-11,本机):该路线的视觉质量显著优于 design_deck/minimal_deck 单独出稿,代价是执行轮次更长——重要汇报优先用此路线,并配合下方的视觉迭代循环。

### ② 快速稿:`scripts/design_deck.py`(84 风格自动版式)
```bash
pip install pptx-designer
python scripts/design_deck.py outline.json -o deck.pptx                    # 84种风格
python scripts/design_deck.py outline.json --template 单位模板.pptx -o deck.pptx  # VI 模板合规
```
- 把 Step 2 大纲 + Step 3 文稿整理成 outline.json(title/subtitle/slides[cover|toc|kpi|bullets|table]);
- 注意「引擎隐性约束」清单(METRICS≤4条、无原生表格等)。

**引擎隐性约束(踩坑实测,违反会静默失败,design_deck.py 已内置规避):**
- goal 页型是故事枚举:hook/problem/solution/features/data/code/cta;hook(封面)**不渲染 bullets**;
- METRICS 页(data)**最多 4 条指标,超出整页静默空白**;键值分隔必须是 `": "`(冒号+空格)才会渲染成"大数值+小标签"卡;
- features 页每页 ≤6 条,超过易溢出;
- `--palette` 只认 COLOR_PALETTES 的 30 个名字(ocean-blue/midnight-navy/cherry-red/wine-burgundy/ink-wash 等),传其他名字**静默回落 ocean-blue**(政务推荐 midnight-navy/cherry-red);
- 引擎**无原生表格**:table 类型只能退化为要点列举,规范表格页(表头/对齐/斑马纹)应改用 `minimal_deck.py` 生成;
- 同一 goal 在不同 layout/decoration 组合下渲染有差异 → **必须走下方迭代循环**,不能生成即交付。

### ③ 次选:`pptx` skill(html2pptx)
环境里有官方 pptx skill 时走其 html2pptx 工作流:HTML 设计遵循 Token 表 + 本 skill 的四步分层法/HSL 配色规则。

### ④ 兜底:自带脚本(任何能跑 Python 的环境)
```bash
pip install python-pptx          # 唯一依赖
python scripts/minimal_deck.py outline.json -o deck.pptx
```
版式支持:cover(封面)/ toc(目录)/ kpi(指标卡)/ bullets(要点)/ table(表格),输出全部为**原生可编辑元素**(无烘焙图片),配色取 Token 表或默认政务红。这是"60分保底",已按本工作流红线实现(每页≤80字、大数字指标卡、状态表)。

### 传给设计引擎的输入清单(缺一不可)
1. 逐页大纲(页码/标题/要点/展示形式);
2. 每页文稿(≤80字/页);
3. 设计 Token 表(或模板文件);
4. 数据文件(台账,供图表取数);
5. 图片素材清单(工作照/搜图,见 charts-and-images.md)。

---

## 三、视觉迭代循环(编排层,Step 4 的灵魂——生成≠交付)

引擎是"自动版式渲染器",一次成稿质量不稳定,**必须走"生成→渲染→审查→调参重生成"循环,最多3轮**:

```
① 生成   design_deck.py outline.json --palette <政务色> --seed N
② 渲染   export_deck.ps1/.sh → 逐页 PNG
③ 审查   三道闸(顺序执行):
         a. verify_deck.py 结构检查(文字越界=FAIL;装饰出血仅WARN)+ OCR 数字对账
         b. agent 逐页看 PNG:内容完整性优先(空页/丢条目 > 配色偏差)、占位符、断行孤字
         c. ZCode 有 visual-judge 则派发逐页判定
④ 定位   对照「引擎隐性约束」清单找根因(如:KPI>4条整页空白=METRICS上限;页面全空=goal映射错)
⑤ 调整   改 outline 条数/映射/分隔符 → 或换 --style/--palette/--seed 重掷 → 重回①
```

判停规则:内容级问题(空页/丢内容/占位符)必须清零;样式级问题(留白比例/断行)2轮内不收敛即接受并告知用户残余项。**规范表格/超4项指标墙等引擎弱项,直接改用 minimal_deck.py 生成该页**,不要硬调引擎。

## 四、自动化验收脚本(循环②③的具体命令)

```bash
python scripts/verify_deck.py deck.pptx              # 越界/空页/每页字量,确定性检查
python scripts/export_deck.ps1 deck.pptx export      # 渲染PNG+PDF(或 mac/Linux 用 export_deck.sh)
python scripts/verify_deck.py deck.pptx --ocr export/png --facts "258,050" "99.46%"   # 数字对账(需rapidocr)
```
