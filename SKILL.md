---
name: docx-smart-format
description: 读取 `.docx` 或从零生成 `.docx`：LLM 判断文档类型、结构与行内格式语义，输出格式决策，本地引擎执行写回。用于自动整理论文/报告/纪要排版、统一字体字号缩进行距分页、处理图片表格公式、配置页眉页脚页码。
---

# DOCX Smart Format

## 平台与运行依赖

- **当前仅支持 Windows x64**：`engine/runtime/docx-auto-template-engine.exe` 是 .NET self-contained 部署，运行时随包分发。macOS / Linux 暂不支持。
- `engine/src` 仅为引擎源码，**运行时不依赖**。重新构建引擎才用得到，正常调用 skill 时忽略它。
- skill 运行时不联网、不依赖外部样式服务或远程模板库，所有规则放在 `references/` 本地。

## 目标

这个 skill 支持两类工作：

1. **模式一**：读取用户上传的 `.docx`，识别文档类型与结构内容，决定调用哪些 Word 格式能力，输出新 `.docx`。
2. **模式二**：不依赖源文档，按 LLM 生成的结构与格式决策从零输出 `.docx`。

重点：

1. LLM 识别文档结构与语义，判断该用什么格式能力，把判断结果映射成稳定、可执行的格式决策。
2. LLM **不直接拼 Word XML**，由本地引擎执行写回。
3. 能用模板解决的优先映射到模板，能用确定性字段表达的不要输出模糊自然语言。
4. 图片、表格、公式、题注、页眉页脚、页码、分节都要作为正式对象处理。
5. 公式是否编号、行内格式是否调整都由 LLM 按语义判断，不做全文逐字硬改。

## 模式判定（机械规则）

```
if 用户提供了源 .docx:
    走模式一：analyze → LLM 生成 decision → apply
elif 用户提供了 主题/提纲/章节结构 + 内容:
    走模式二：LLM 生成 RenderSpec → render
else:
    向用户澄清缺哪类输入，不要猜
```

### 模式一：基于现有 docx 重排

```powershell
<skill-dir>\engine\runtime\docx-auto-template-engine.exe analyze --input <input.docx> --output <analysis.json>
<skill-dir>\engine\runtime\docx-auto-template-engine.exe apply --source <source.docx> --decision <decision.json> --output <result.docx> --template <template.docx>
```

适用：用户上传现有 Word 文档、需保留原内容/图表/公式与行内样式语义、统一格式。

### 模式二：直接生成新 docx

```powershell
<skill-dir>\engine\runtime\docx-auto-template-engine.exe render --spec <render-spec.json> --output <result.docx> --template <template.docx>
<skill-dir>\engine\runtime\docx-auto-template-engine.exe render --spec <render-spec.json> --output <result.docx> --template-preset builtin-undergraduate-thesis
```

适用：用户只给主题、提纲、模板要求或内容片段，目标是生成新的正式文档。

毕业论文场景下，`--template-preset builtin-undergraduate-thesis` 自动启用 `--normalize-references`。如需强制开/关，显式传 `--normalize-references` 或 `--normalize-references false`。命令行参数 `=` 与空格两种写法（`--template-preset=...` / `--template-preset ...`）均接受，文档统一用空格形式。

## 生成前用户确认（必须执行）

无论模式一还是模式二，**开始生成之前必须主动询问用户字体偏好**：

> 生成前确认一下字体设置，推荐默认：中文宋体、英文 Times New Roman、正文小四。您有特别要求吗？

默认值见 [references/template-catalog.md](references/template-catalog.md) 的"通用硬规则"。

满足以下任一条件可跳过询问：

- 用户已在原始要求中明确说明了字体/字号
- 用户明确表示"用默认就行"或等价说法
- 模式一中用户回答"保持原样"，则不覆盖原文档字体

确认后，把用户选择写入 `RenderSpec.document` 以及所有正文段落的 `run.eastAsiaFont` / `run.asciiFont` / `run.fontSize`，确保全篇一致。

## LLM 要做的判断

### 1. 文档类型

论文、实验报告、公司汇报、会议纪要、通用正式文档。

### 2. 块级结构

文档标题 / 一二三级标题 / 正文 / 列表 / 引用 / 图注 / 表注 / 公式 / 参考文献 / 附录。

### 3. 页面与分节结构

是否分节、节间页码格式是否不同、首页是否单独、是否奇偶页不同、章节是否新页起。

### 4. 行内格式语义

哪些词加粗、哪些英文术语斜体、哪些字符上标/下标、哪些位置保留下划线、哪些地方需要 Tab 对齐。

### 5. 公式与对象关系

行内/独立、是否编号、编号是否被正文引用、图注属于哪张图、表注属于哪个表。

## 何时调用何种格式能力

### 分节
封面、摘要、目录、正文、附录之间存在页码/页眉页脚/方向/边距差异时分节。不要只因为视觉换页就新建分节。

### 字体与字号
默认值与标题层级字号对应表，见 [references/template-catalog.md](references/template-catalog.md) 的"通用硬规则"。中英文混排必须分别设置中文字体和英文字体，不要只设一个总字体。

### 加粗
内容属于标题、强调词、术语、表头、关键结论时使用。

### 斜体
内容属于英文术语、变量名、书名式强调、学术记号时使用。

### 上标 / 下标
先判断语义再改样式：
- `H2O`、`x^2`、`x_1`、`CO2`、`m/s^2` 按下标/上标处理
- 变量幂、下标、化学计量数、脚注标记优先用 `superscript` / `subscript`
- 正文里的强调数字、编号、页码、标题序号**不要**误判成上标/下标
- 不确定时回退 `baseline`

### 下划线
模板或语义明确要求强调线、签名线、填空线、术语标识时使用。

### Tab
左右对齐、目录页码对齐、编号后文本对齐、页眉页脚布局时使用。**不要用 Tab 代替首行缩进**。

### 首行缩进
正文段、说明段、一般叙述段且模板要求缩进时使用。标题、图注、表注、列表项、独立公式通常不使用。

### 悬挂缩进
参考文献、编号列表、多行列表、条文型结构时使用。

### 行间距 / 段前段后
正文、标题、图注、表注、参考文献需要不同行距时调用；标题层级、图注表注、公式前后需要视觉分隔时调用段前段后。**不要靠空行模拟段前段后**。

### 页眉页脚
存在首页不同、奇偶页不同、章节页眉、规范页脚信息时使用。页眉页脚按节处理，不假设全篇只有一套。

**奇偶页眉、STYLEREF 字段、apply 模式 r:id 盘点等深度规则见 [references/header-footer-odd-even.md](references/header-footer-odd-even.md)。**

### 页码
正文从指定页开始编号、前置部分不编号或编号规则不同、页码位置有要求时使用。**不要用普通数字代替页码字段**。

**页码涉及 `pgNumType`（节级编号格式）与 `PAGE` 字段（页面显示）两层，必须同时设置。详见 [references/header-footer-odd-even.md](references/header-footer-odd-even.md) 的"页码"章节。**

### 分页控制
一级标题前换页、封面与正文分开、标题不能落在页尾时使用。优先用分页属性，**不用空段落顶页**。

### 图片
居中、宽高控制、保持比例、图注位置、图题样式、图片前后间距。
- 普通插图：版心 `60%` 到 `80%`
- 宽图或流程图：`85%` 到 `100%`
- 多张同类图：尽量统一宽度
- 默认居中，图注默认放图下方
- 公式截图按图片处理，不伪装成可编辑公式

### 表格
表头样式、列宽、边框、单元格对齐、表注位置、表格前后间距。
- 默认三线表（顶线、表头下横线、底线），不使用满框边线
- 数据表：表头加粗，数值列居中或右对齐
- 整体宽度 `80%` 到 `100%`，默认居中
- 表注默认放表上方，表内一般不首行缩进

### 公式
- 行内公式还是独立公式、是否需要编号、编号是否右侧、是否需要说明文字、前后间距和对齐
- **是否编号必须由 LLM 按文档类型/模板规范/上下文引用关系判断，不是所有公式都要编号**
- 正文中以"式(1)""公式(2)"引用的通常编号；普通推导短公式、临时计算式通常不编号
- 行内公式通常不编号；独立公式也不默认强制编号

### 参考文献与正文交叉引用
毕业论文 / 正式技术报告里，本地引擎提供**参考文献规范化**能力（书签 + REF 字段交叉引用 + `[n]<TAB>` 条目规范化）。

**完整规则、开关方式、约定边界、自检要求见 [references/reference-normalization.md](references/reference-normalization.md)。**

LLM 只需决定"哪些 `[n]` 是上标"，书签 / 域 / Tab / 悬挂缩进由引擎统一执行。

### 横向页面 + 竖向页码文本框
含横向插页（宽表、宽图、流程图）且页码需竖向排在左侧装订边时使用。

**完整 RenderSpec 字段、OOXML 结构、wps+VML 双轨说明见 [references/landscape-vertical-pagenum.md](references/landscape-vertical-pagenum.md)。**

## 决策输出层次

LLM 输出**决策**，不是解释长文。最少分四层：

1. **文档级规则**：页面大小、页边距、字体、基础字号、行距、页眉页脚策略、页码策略
2. **块级决策**：每段是 `heading1` / `body` / `reference` 等；是否分页前置、首行缩进、题注样式
3. **行内决策**：每个 run 是否加粗 / 斜体 / 上下标 / Tab
4. **对象级决策**：图片宽度与图注位置、表格宽度与表注位置、公式对齐与是否编号

完整字段见 [references/decision-schema.md](references/decision-schema.md)。

## 本地引擎可执行能力

LLM 只需输出决策，以下能力已经在本地引擎落地：

- **段落级**：对齐、首行缩进、悬挂缩进、左右缩进、段前段后、行距与行距规则、分页前
- **run 级**：中文字体、英文字体、字号、加粗、斜体、下划线、字体颜色、上标、下标、文本替换、Tab 保留与写回
- **文档级**：分节、页面方向、页边距、页眉页脚、页码字段
- **对象级**：图片写回、表格写回、原生公式写回、页眉/页脚锚定文本框（`wps:wsp` + VML `v:rect` 双输出，支持 `vert270` 竖排页码）、参考文献条目书签（`_Ref_ref_<n>`）+ 正文上标 REF 字段交叉引用

能力边界与调用说明见 [references/local-capability-map.md](references/local-capability-map.md)。

## 禁止做法（集中清单）

- 不要让 LLM 直接输出 Word XML
- 不要用空行模拟段前段后
- 不要用普通文本数字代替页码字段
- 不要让 LLM 直接写 "I" / "II" / "1" 等文本数字代替 `PAGE` 字段
- 不要只设 `pgNumType` 不写 `PAGE` 字段
- 不要让正文沿用前置节的 footer 文件
- 不要用 Tab 代替所有缩进（特别是首行缩进）
- 不要把所有粗体、斜体都原样保留——必须先判断语义
- 不要把图片公式伪装成原生公式
- 不要把所有独立公式一律自动编号
- STYLEREF 字段不要使用 `\s` 开关——使用 `\* MERGEFORMAT`
- apply 模式下不要按 `r:id` 顺序或文件名顺序判断页眉绑定——按 `w:type`
- 横向竖排页码不要用空白字符 + 旋转字体伪造，必须走 `wps:wsp` / VML `v:rect` 文本框
- 横向页文本框不要做成 inline drawing——必须 anchored + wrapNone

## 失败回退

遇到不确定时采用保守策略：

- 文档类型不确定：回退 `general-formal-doc`
- 段落角色不确定：回退 `body`
- 行内样式不确定：保留原样式
- 图注表注归属不确定：保留相邻顺序
- 公式编号归属不确定：保留原顺序
- 是否该分节不确定：不新建分节
- 是否该编号不确定：不编号，除非模板明确强制

## 故障排查

| 现象 | 可能原因 | 处置 |
|------|---------|------|
| 偶数页页眉空白 | STYLEREF 用了 `\s` 开关，但一级标题不是 Word 自动编号 | 把模板里 STYLEREF 改成 `\* MERGEFORMAT` |
| 前置部分莫名显示"摘要"等文字 | 模板里 header4.xml 硬编码了文字，前置节没显式 clear | 在前置节 `headerFooterDecisions` 显式 `action: "clear"` |
| 奇偶页页眉一样 | `oddEvenDifferent` 未设 `true` 或只绑定了 default 一条引用 | 同时设置 `oddEvenDifferent: true` 并在 sectionDecisions 用 `headerDecisionKeys` 绑定 default + even |
| 前置部分根本不显示页码 | 设了 `pgNumType` 但 footer 没含 `PAGE` 字段 | footer 加 `pageNumberField: true`，或前置不显示页码时连 `pgNumType` 也别写 |
| 正文页码从大数字开始而不是 1 | 正文沿用了前置节的 footer 文件 | 正文起始节换一个独立 footer 文件 |
| 横向页竖排页码定位漂移 | 文本框 `relativeFrom` 写成了 `margin` 而非 `page` | 改回 `relativeFromH: "page"` / `relativeFromV: "page"` |
| 引擎报"找不到 styleId"（apply 模式） | 决策里引用了模板没有的样式 key | 检查 `--template` 是否传对，或换用 `general-formal-doc` 模板 |
| `analyze` 输出里 sectPr header 引用为空 | 源 docx 的节没有任何页眉页脚 | LLM 在决策里要主动声明 header/footer，不能假设有继承 |

引擎报错优先读取 stderr 内容；常见错误码对照见 `engine/runtime/docx-auto-template-engine.exe --help`（如有）。

## 资源索引

- [references/decision-schema.md](references/decision-schema.md) — 决策 JSON 的字段定义与格式能力映射
- [references/template-catalog.md](references/template-catalog.md) — 内置模板与默认字体字号规则（**字体默认值唯一权威**）
- [references/local-capability-map.md](references/local-capability-map.md) — 本地分析/写回能力清单与 LLM 调用边界
- [references/header-footer-odd-even.md](references/header-footer-odd-even.md) — 奇偶页眉 / STYLEREF / pgNumType / apply r:id 盘点
- [references/reference-normalization.md](references/reference-normalization.md) — 参考文献条目书签 + 正文 REF 字段交叉引用
- [references/landscape-vertical-pagenum.md](references/landscape-vertical-pagenum.md) — 横向页面左侧竖排页码（wps + VML 双轨）
- [scripts/sample-decision.json](scripts/sample-decision.json) — 模式一决策样例
- [scripts/sample-render-spec.json](scripts/sample-render-spec.json) — 模式二 RenderSpec 样例
- [scripts/sample-chem-render-spec.json](scripts/sample-chem-render-spec.json) — 化学式 / 单位指数样例
- [scripts/validate_decision.py](scripts/validate_decision.py) — 决策 JSON 校验器
