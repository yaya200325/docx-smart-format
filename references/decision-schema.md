# Decision Schema

## 目标

该文档定义 `docx-smart-format` skill 中 LLM 应输出的决策结构。

重点是表达“该调用哪些 Word 格式能力”，不是表达底层实现。

## 顶层结构

```json
{
  "docType": "academic-thesis",
  "templateId": "academic-thesis",
  "confidence": 0.92,
  "documentRules": {},
  "headerFooterDecisions": [],
  "sectionDecisions": [],
  "blockDecisions": [],
  "runDecisions": [],
  "equationDecisions": [],
  "assetDecisions": []
}
```

## 文档级规则 `documentRules`

用于控制整篇文档或整节文档的格式能力。

示例：

```json
{
  "pageSize": "a4",
  "orientation": "portrait",
  "marginTop": "1440",
  "marginBottom": "1440",
  "marginLeft": "1800",
  "marginRight": "1800",
  "cnFontFamily": "宋体",
  "enFontFamily": "Times New Roman",
  "baseFontSizePt": 12,
  "lineSpacing": 1.5,
  "headerPolicy": "template",
  "footerPolicy": "template",
  "pageNumberPolicy": "body-start",
  "firstPageDifferent": true,
  "oddEvenDifferent": false
}
```

字段说明：

- `oddEvenDifferent`：是否启用奇偶页眉页脚。设为 `true` 会触发 settings.xml 写入 `<w:evenAndOddHeaders/>`，并要求每个相关节同时提供 `default` 与 `even` 两条引用。**只要 LLM 决定让奇偶页眉显示不同内容，就必须设为 `true`，否则 Word 渲染时奇偶页都会落到同一份页眉上。**
- `firstPageDifferent`：是否启用首页不同。设为 `true` 时该节需要再提供 `first` 类型的引用。
- `headerPolicy` / `footerPolicy`：声明这一节页眉页脚的来源。常用值：`template`（继承模板）、`override`（由 `headerFooterDecisions` 显式覆盖）、`clear`（清空）。
- `pageNumberPolicy`：声明页码策略。常用值：`none`、`continuous`、`body-start`（前置不计页码、正文开始新编号）、`roman-then-arabic`（前置罗马数字、正文阿拉伯数字）。具体落点见 `sectionDecisions` 的 `pgNumType` 字段。

## 页眉页脚决策 `headerFooterDecisions`

用于声明每一份页眉/页脚 XML 部件的内容与角色。**每条决策只描述一份部件，节与部件的绑定关系由 `sectionDecisions` 完成。** 这样可以在一节里同时绑定 `default` + `even` 两条引用，从而真正实现奇偶页眉。

字段：

- `decisionKey`：本条决策的稳定 ID，供 `sectionDecisions` 引用。
- `kind`：`header` 或 `footer`。
- `type`：OOXML 中的 `w:type`，取值 `default` / `even` / `first`。
  - 注意：`odd` 不是合法 OOXML 值。如果 LLM 想表达"奇数页"，请用 `default`，并同时把 `documentRules.oddEvenDifferent` 设为 `true`。
- `action`：本部件的处理动作。常用值：
  - `useTemplate`：直接沿用模板里这份页眉页脚。
  - `clear`：清空内容（前置部分常用）。
  - `replace`：使用 `sourcePath` 指向的新 XML 重写。
- `sourcePath`（可选）：当 `action = "replace"` 时使用，指向新的 header/footer XML 文件。
- `applyStyleKey`（可选）：写回时套用的样式。
- `targetSectionKey`（可选）：声明这份部件天然属于哪个节，便于 apply 模式自检。

奇偶页眉的最小决策结构示例（正文节：奇数页显示论文标题，偶数页显示 STYLEREF 取得的一级标题）：

```json
{
  "documentRules": {
    "oddEvenDifferent": true,
    "headerPolicy": "override",
    "pageNumberPolicy": "roman-then-arabic"
  },
  "headerFooterDecisions": [
    {
      "decisionKey": "frontMatter.header.empty",
      "kind": "header",
      "type": "default",
      "action": "clear"
    },
    {
      "decisionKey": "frontMatter.header.empty.even",
      "kind": "header",
      "type": "even",
      "action": "clear"
    },
    {
      "decisionKey": "body.header.title",
      "kind": "header",
      "type": "default",
      "action": "replace",
      "sourcePath": "header-body-odd.xml"
    },
    {
      "decisionKey": "body.header.styleRef",
      "kind": "header",
      "type": "even",
      "action": "replace",
      "sourcePath": "header-body-even.xml"
    },
    {
      "decisionKey": "frontMatter.footer.roman",
      "kind": "footer",
      "type": "default",
      "action": "replace",
      "sourcePath": "footer-front-roman.xml"
    },
    {
      "decisionKey": "body.footer.page",
      "kind": "footer",
      "type": "default",
      "action": "replace",
      "sourcePath": "footer-body-page.xml"
    }
  ]
}
```

要点：

- 奇数页内容写到 `type = "default"` 的部件，偶数页内容写到 `type = "even"` 的部件。
- 即使奇/偶有一边为空，也必须显式提供 `clear` 决策，否则 Word 会回退到唯一一份 default，奇偶页都会显示。
- `sourcePath` 指向的 XML 必须是合法的 `<w:hdr>` / `<w:ftr>` 根。STYLEREF 字段必须由 LLM 在该 XML 中明确写出 `<w:fldChar w:fldCharType="begin"/>` … `<w:instrText xml:space="preserve"> STYLEREF "Heading 1" \* MERGEFORMAT </w:instrText>` …，引擎不会自动生成。
- STYLEREF 的 switch 用 `\* MERGEFORMAT`，**不要用 `\s`**。`\s` 取段落编号，对手动编号的章节会返回空。

## 节决策 `sectionDecisions`

每条节决策描述"在某个块之后开始一个新节"，并把该节绑定到具体的页眉页脚部件。

字段：

- `afterBlockId` / `afterBlockPath`：节边界落点。
- `sectionType`：`continuous` / `nextPage` / `oddPage` / `evenPage`。
- `headerDecisionKey` / `footerDecisionKey`：单条引用。如果该节只需要一份 default 页眉页脚，用这一对即可。
- `headerDecisionKeys` / `footerDecisionKeys`：**多条引用**。需要同时绑定 default + even（或再加 first）时，必须用复数字段，把多个 `decisionKey` 全部列出。引擎按 `type` 去重，不按出现顺序。
- `pgNumType`：本节页码字段。
  - `format`：`decimal` / `upperRoman` / `lowerRoman` / `upperLetter` / `lowerLetter`。
  - `start`：起始页码整数；省略表示沿用上一节。

示例：前置部分罗马数字、正文阿拉伯数字、正文奇偶页眉不同。

```json
{
  "sectionDecisions": [
    {
      "afterBlockId": "frontMatter.end",
      "sectionType": "nextPage",
      "headerDecisionKeys": [
        "frontMatter.header.empty",
        "frontMatter.header.empty.even"
      ],
      "footerDecisionKey": "frontMatter.footer.roman",
      "pgNumType": {
        "format": "upperRoman",
        "start": 1
      }
    },
    {
      "afterBlockId": "body.start",
      "sectionType": "nextPage",
      "headerDecisionKeys": [
        "body.header.title",
        "body.header.styleRef"
      ],
      "footerDecisionKey": "body.footer.page",
      "pgNumType": {
        "format": "decimal",
        "start": 1
      }
    }
  ]
}
```

要点：

- 同一个 `sectionDecisions` 节里出现 `default` + `even` 时，引擎会在该节 sectPr 写入两条 `<w:headerReference>`，并自动确保 settings.xml 里有 `<w:evenAndOddHeaders/>`。
- 如果只写了一条（比如只列了 `default`），即使 `oddEvenDifferent = true`，奇偶页眉也不会真正分开。
- `pgNumType.start` 只在该节生效。前置部分若想从 I 起，正文若想从 1 起，必须分别声明。
- 不要把页码以普通数字写到段落里。页码必须经由 `PAGE` 字段，写在 footer XML 内；起始值由 `pgNumType.start` 决定。

## apply 模式额外约束

LLM 在 apply 模式下还需要：

- 先盘点源 docx 里所有 `sectPr` 已有的 `headerReference` / `footerReference`，按 `(节, type)` 二元组列清单。
- 任何**未被本次决策覆盖**的 `type`，要么沿用模板内容，要么显式 `clear`。**不要假设引擎会自动清空遗留页眉**——上一版本就是因为忽略了模板里残留的 `"摘要"` header4，导致前置部分仍然显示"摘要"。
- 写回完成后，再从 sectPr 视角自检一次：每个节是否拿到了它应有的 default/even/first 全部引用，r:id 是否指向决策声明的那份部件。**不要靠 r:id 数字顺序或文件名顺序判断绑定关系，永远按 `w:type` 判断。**

## 块级决策 `blockDecisions`

用于控制标题、正文、图注、表注、列表、参考文献等块级内容。

示例：

```json
{
  "path": "/body/paragraph[12]",
  "kind": "paragraph",
  "semanticRole": "heading2",
  "targetStyleKey": "template.heading2",
  "pageBreakBefore": false,
  "keepWithNext": true,
  "firstLineIndentChars": 0,
  "hangingIndentChars": 0,
  "align": "left",
  "lineSpacing": 1.5,
  "beforeSpacingPt": 12,
  "afterSpacingPt": 6,
  "useTabStops": false
}
```

## 行内决策 `runDecisions`

用于控制加粗、斜体、下划线、颜色、上标、下标、Tab 等局部格式。

示例：

```json
{
  "runPath": "/body/paragraph[8]/run[3]",
  "preserve": true,
  "bold": false,
  "italic": true,
  "underline": null,
  "fontColor": null,
  "verticalAlign": "baseline",
  "containsTab": false
}
```

## 公式决策 `equationDecisions`

用于控制原生公式对象。

示例：

```json
{
  "path": "/body/equation[12.1]",
  "kind": "equation",
  "displayMode": "display",
  "semanticRole": "equation",
  "targetStyleKey": "template.equation",
  "align": "center",
  "beforeSpacingPt": 6,
  "afterSpacingPt": 6,
  "numbering": {
    "enabled": true,
    "format": "(1)",
    "position": "right"
  }
}
```

规则：

- `numbering.enabled` 不能默认写成 `true`。
- 是否编号必须由 LLM 根据模板、文档类型、正文引用关系判断。
- 行内公式通常不编号。

## 对象决策 `assetDecisions`

用于控制图片和表格。

### 图片示例

```json
{
  "path": "/body/image[3]",
  "kind": "image",
  "align": "center",
  "maxWidthPct": 80,
  "keepAspectRatio": true,
  "spaceBeforePt": 6,
  "spaceAfterPt": 6,
  "captionPath": "/body/paragraph[19]",
  "captionPosition": "below"
}
```

### 表格示例

```json
{
  "path": "/body/table[2]",
  "kind": "table",
  "headerRow": true,
  "align": "center",
  "preferredWidthPct": 100,
  "borderStyle": "single",
  "cellAlign": "center",
  "spaceBeforePt": 6,
  "spaceAfterPt": 6,
  "captionPath": "/body/paragraph[22]",
  "captionPosition": "above"
}
```

## 角色集合

- `title`
- `heading1`
- `heading2`
- `heading3`
- `body`
- `bullet`
- `numbered`
- `quote`
- `figureCaption`
- `tableCaption`
- `equation`
- `appendix`
- `reference`

## 调用约束

- 需要页码时，输出页码策略，不输出普通数字文本。
- 需要页眉页脚时，输出页眉页脚策略。**奇偶页眉必须同时设置 `oddEvenDifferent = true` 并在 `sectionDecisions` 里同时绑定 default + even 两条引用，缺一不可。**
- 需要 STYLEREF 取当前章节标题时，使用 `\* MERGEFORMAT` 开关，禁止使用 `\s`。
- 需要首行缩进时，输出缩进字段，不用 Tab 假装。
- 需要列表或参考文献排版时，优先考虑悬挂缩进。
- 需要化学式和数学角标时，优先使用上标/下标字段。
- 需要图注表注时，输出对象关系与位置。
- 公式是否编号必须交给 LLM 判断，不做默认强制。

## 回退规则

- `docType` 不确定时，用 `general-formal-doc`
- `semanticRole` 不确定时，用 `body`
- `verticalAlign` 不确定时，用 `baseline`
- 是否分节不确定时，优先不分节
- 图片公式不输出到 `equationDecisions`
- 图注表注归属不确定时，保留相邻顺序
- 公式是否编号不确定时，优先 `numbering.enabled = false`
