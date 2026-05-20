# OMML Cookbook（render 模式公式写法）

## 这份文档的作用

render 模式下从零生成公式时，`equation` 块的 `xml` 字段要求 LLM 直接提交合法 OMML 片段。这是 LLM 在本 skill 中**唯一**被允许直接输出 Word XML 的地方。

本文给出：
- equation block 字段定义
- OMML 命名空间硬规则
- 何时走 equation 块、何时不走
- 各种数学结构的最小可粘贴 OMML 模板
- displayMode 行为说明
- 自检清单

apply 模式公式不在本文范围（OMML 透传不改，节奏交给 `equationDecisions`，见 [decision-schema.md](decision-schema.md)）。

## equation block 字段定义

```jsonc
{
  "kind": "equation",
  "equation": {
    "displayMode": "display",  // "display"（独立公式，居中）或 "inline"（行内）
    "xml": "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">...</m:oMathPara>",
    "text": "可选：纯文本回退，仅在 xml 解析失败时显示"
  }
}
```

字段语义：

| 字段 | 必填 | 说明 |
|---|---|---|
| `xml` | **强烈推荐** | 完整 OMML 片段，最外层是 `<m:oMathPara>` 或裸 `<m:oMath>`，**最外层必须自带 `xmlns:m` 命名空间声明** |
| `displayMode` | 必填 | `display` 独立成段、居中；`inline` 跟随段落 |
| `text` | 可选 | 仅 `xml` 为空或解析失败时显示，作为兜底 |

## OMML 命名空间硬规则

**最外层 OMML 元素必须显式声明命名空间**，否则引擎反序列化会失败回退为占位文本。

```xml
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  ...
</m:oMathPara>
```

要点：

- 命名空间 URL **完全等于**：`http://schemas.openxmlformats.org/officeDocument/2006/math`
- 所有 OMML 元素必须带 `m:` 前缀
- 内层元素不需要重复声明命名空间
- `<m:oMathPara>` 用于 `displayMode:"display"`；`<m:oMath>` 用于 `displayMode:"inline"`（也可独立用，本文统一用 `<m:oMathPara><m:oMath>...</m:oMath></m:oMathPara>` 包一层）

## 何时走 equation 块

| 内容形态 | 走 equation 块 | 走 run 上下标 | 走普通 `run.text` |
|---|---|---|---|
| 数学表达式（等式 / 不等式 / 推导 / 计算式） | ✅ | ❌ | ❌ |
| 数学变量（正文中的 x、y、n、α…） | ✅ | ❌ | ❌ |
| 物理单位含指数 / 复合（m/s²、A·m⁻¹、kg·m·s⁻²） | ✅ | ❌ | ❌ |
| 化学式（CO₂、H₂O、Fe³⁺） | ❌ | ✅ | ❌ |
| 化学方程式（2H₂+O₂ → 2H₂O） | ❌ | ✅ | ❌ |
| 简单单位（kg、m、s、Pa、°C） | ❌ | ❌ | ✅ |
| 脚注标号 / 页码 / 标题序号 | ❌ | 仅脚注用 ✅ | ✅ |
| 公式截图 | 当图片处理（`kind:"image"`） | ❌ | ❌ |

关键边界：

- **变量必走 equation 块**。哪怕只是"设 *x* 为..."里的单字母 `x`，也用 `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>` 包一下，保证 Word 中字体（Cambria Math）和斜体（自动）正确。
- **化学式不走 equation**。原因：化学式没有数学变量语义，OMML 会把所有字母强制斜体化，效果反而错。继续用 `run.superscript` / `run.subscript`。
- **公式截图** 仍是图片，不要伪造成 OMML。

## OMML cookbook（最小可粘贴模板）

下面所有片段省略外层 `<m:oMathPara xmlns:m="...">` 与 `</m:oMathPara>`，使用时**必须**自己补上最外层带命名空间的包装。

### 1. 纯文本与符号

```xml
<m:oMath>
  <m:r><m:t>E=mc^2</m:t></m:r>
</m:oMath>
```

要点：
- 文本必须放在 `<m:r><m:t>...</m:t></m:r>`，**不能裸放在 `<m:oMath>` 下**
- 希腊字母、特殊符号直接用 Unicode 字符放入 `<m:t>`：`α β γ δ ε ζ η θ λ μ π ρ σ τ φ ψ ω Σ Π Ω Δ Γ Φ ∞ ∂ ∇ ∫ ∮ ∑ ∏ √ ± × ÷ ≤ ≥ ≠ ≈ ≡ ∈ ∉ ⊂ ⊃ ∪ ∩ →`

### 2. 上标 `<m:sSup>`

`x²`：

```xml
<m:sSup>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
</m:sSup>
```

### 3. 下标 `<m:sSub>`

`x₁`：

```xml
<m:sSub>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
  <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
</m:sSub>
```

### 4. 上下标合用 `<m:sSubSup>`

`x_i^2`：

```xml
<m:sSubSup>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
  <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
  <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
</m:sSubSup>
```

### 5. 分式 `<m:f>`

`a / b`（默认横线分式）：

```xml
<m:f>
  <m:num><m:r><m:t>a</m:t></m:r></m:num>
  <m:den><m:r><m:t>b</m:t></m:r></m:den>
</m:f>
```

分式类型（`<m:fPr><m:type m:val="..."/>`）：

| `m:val` | 含义 |
|---|---|
| `bar` | 横线分式（默认，等同省略） |
| `noBar` | 无分线（二项式） |
| `skw` | 斜线分式（`a/b` 倾斜） |
| `lin` | 线性分式（行内 `a/b`） |

### 6. 根号 `<m:rad>`

`√x`（平方根，隐藏次数）：

```xml
<m:rad>
  <m:radPr><m:degHide m:val="1"/></m:radPr>
  <m:deg/>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
</m:rad>
```

`³√x`（三次方根）：

```xml
<m:rad>
  <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
</m:rad>
```

要点：
- 平方根时 `<m:deg/>` 必须空，并加 `<m:radPr><m:degHide m:val="1"/></m:radPr>`
- 任意次方根时 `<m:deg>` 内填写次数

### 7. 大型算子 `<m:nary>`（求和 / 积分 / 连乘）

`∑_{n=1}^{∞} aₙ`：

```xml
<m:nary>
  <m:naryPr><m:chr m:val="∑"/></m:naryPr>
  <m:sub><m:r><m:t>n=1</m:t></m:r></m:sub>
  <m:sup><m:r><m:t>∞</m:t></m:r></m:sup>
  <m:e>
    <m:sSub>
      <m:e><m:r><m:t>a</m:t></m:r></m:e>
      <m:sub><m:r><m:t>n</m:t></m:r></m:sub>
    </m:sSub>
  </m:e>
</m:nary>
```

`∫_a^b f(x) dx`：

```xml
<m:nary>
  <m:naryPr><m:chr m:val="∫"/></m:naryPr>
  <m:sub><m:r><m:t>a</m:t></m:r></m:sub>
  <m:sup><m:r><m:t>b</m:t></m:r></m:sup>
  <m:e>
    <m:r><m:t>f(x)dx</m:t></m:r>
  </m:e>
</m:nary>
```

要点：
- `<m:chr m:val="..."/>` 取值常见：`∑`（求和） / `∫`（积分） / `∏`（连乘） / `∮`（环路积分） / `∐`（余积）
- 不写 `<m:chr>` 时默认 `∫`
- 上下限可省略；若同时省略上下限，结构仍合法

### 8. 矩阵 `<m:m>`

`[[a, b], [c, d]]`：

```xml
<m:d>
  <m:dPr>
    <m:begChr m:val="["/>
    <m:endChr m:val="]"/>
  </m:dPr>
  <m:e>
    <m:m>
      <m:mr>
        <m:e><m:r><m:t>a</m:t></m:r></m:e>
        <m:e><m:r><m:t>b</m:t></m:r></m:e>
      </m:mr>
      <m:mr>
        <m:e><m:r><m:t>c</m:t></m:r></m:e>
        <m:e><m:r><m:t>d</m:t></m:r></m:e>
      </m:mr>
    </m:m>
  </m:d>
</m:d>
```

要点：
- `<m:m>` = 矩阵；`<m:mr>` = 行；`<m:e>` = 单元
- 外层用 `<m:d>` 包括号；不写 `<m:dPr>` 时默认圆括号
- 行列式用 `<m:begChr m:val="|"/>`+`<m:endChr m:val="|"/>`

### 9. 括号 `<m:d>`

`(a + b)`：

```xml
<m:d>
  <m:e>
    <m:r><m:t>a+b</m:t></m:r>
  </m:e>
</m:d>
```

自定义括号：`<m:dPr><m:begChr m:val="{"/><m:endChr m:val="}"/></m:dPr>`（花括号）；`m:val="|"` 绝对值；`m:val="⌊"` / `m:val="⌋"` 取整。

### 10. 多行对齐 `<m:eqArr>`

```
a = b + c
  = d + e
  = f
```

```xml
<m:eqArr>
  <m:e>
    <m:r><m:t>a=b+c</m:t></m:r>
  </m:e>
  <m:e>
    <m:r><m:t>=d+e</m:t></m:r>
  </m:e>
  <m:e>
    <m:r><m:t>=f</m:t></m:r>
  </m:e>
</m:eqArr>
```

要点：用 `=` 前的字符位置自动对齐；每行包 `<m:e>` 一组。

### 11. 函数名 `<m:func>`

`sin x`：

```xml
<m:func>
  <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
</m:func>
```

`log₂ n`：

```xml
<m:func>
  <m:fName>
    <m:sSub>
      <m:e><m:r><m:t>log</m:t></m:r></m:e>
      <m:sub><m:r><m:t>2</m:t></m:r></m:sub>
    </m:sSub>
  </m:fName>
  <m:e><m:r><m:t>n</m:t></m:r></m:e>
</m:func>
```

适用：`sin` `cos` `tan` `cot` `sec` `csc` `log` `ln` `exp` `lim` `max` `min` `sup` `inf` `det` 等。函数名走 `<m:func>` 才能保证字体直立（不变斜体）。

### 12. 上划线 / 下划线 / 向量

`x̄`（上横）：

```xml
<m:bar>
  <m:barPr><m:pos m:val="top"/></m:barPr>
  <m:e><m:r><m:t>x</m:t></m:r></m:e>
</m:bar>
```

`v⃗`（向量、上箭头）：

```xml
<m:acc>
  <m:accPr><m:chr m:val="⃗"/></m:accPr>
  <m:e><m:r><m:t>v</m:t></m:r></m:e>
</m:acc>
```

### 13. 行内单变量

正文 "设 *x* 为..." 中的 `x`，直接：

```xml
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:r><m:t>x</m:t></m:r>
</m:oMath>
```

放在 `equation` block 里，`displayMode:"inline"`。Word 自动用 Cambria Math 斜体渲染。

## displayMode 行为说明

| `displayMode` | 段落效果 | 引擎行为（DocxRenderer.cs:1182-1184） |
|---|---|---|
| `"display"` | 独立成段、自动居中 | 强制 `<w:jc w:val="center"/>`；忽略 `paragraph.lineSpacing`（保留段前段后），防止积分/分式被截断 |
| `"inline"` | 跟随当前段落 | 按 `paragraph.align` 对齐，行距正常 |

**段落对齐 / 段前段后**：通过 `equation.paragraph` 字段控制（与 run 段落格式同结构），无需写进 OMML。

## 自检清单（提交前必过）

- [ ] 最外层 OMML 元素带 `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`
- [ ] 所有 OMML 元素都带 `m:` 前缀
- [ ] 文本都在 `<m:r><m:t>...</m:t></m:r>`，没有裸文本节点
- [ ] `<m:sSup>` / `<m:sSub>` / `<m:sSubSup>` 的 `<m:e>` 不为空
- [ ] `<m:rad>` 平方根时 `<m:deg/>` 为空且配 `<m:degHide m:val="1"/>`
- [ ] `<m:f>` 同时有 `<m:num>` 和 `<m:den>`
- [ ] `<m:m>` 每行 `<m:mr>` 的 `<m:e>` 数量一致
- [ ] `displayMode` 与场景匹配：独立公式用 `display`，正文中的变量 / 短表达式用 `inline`
- [ ] JSON 转义：`equation.xml` 字段里的 `"` 必须 `\"`，反斜杠必须 `\\`

## 本地 dry-run 验证

提交 RenderSpec 给引擎前可本地验证 OMML 是否落地：

```powershell
& "$skill\engine\runtime\docx-auto-template-engine.exe" render `
    --spec ./render-spec.json `
    --output ./result.docx

# 解压 docx 检查 document.xml 含 oMath 节点
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead("./result.docx")
$entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
(New-Object System.IO.StreamReader($entry.Open())).ReadToEnd() | Select-String "oMath"
```

如果 `document.xml` 里看不到 `<m:oMath>`，说明 LLM 输出的 OMML 被引擎解析失败回退为占位文本——回到自检清单逐项检查。

## 反例（常见错误）

```jsonc
// ❌ 把数学表达式塞进 run.text
{ "kind":"body", "runs":[{ "text":"x² + y² = r²" }] }

// ❌ OMML 缺命名空间
{ "kind":"equation", "equation":{ "xml":"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>" } }

// ❌ 裸文本不包 <m:r><m:t>
{ "kind":"equation", "equation":{ "xml":"<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">x=1</m:oMath>" } }

// ❌ 化学式 H₂O 误走 equation（OMML 会把 H 强制斜体化）
{ "kind":"equation", "equation":{ "xml":"<m:oMath xmlns:m=\"...\"><m:r><m:t>H</m:t></m:r><m:sSub>...</m:sSub><m:r><m:t>O</m:t></m:r></m:oMath>" } }
// ✅ 化学式正确做法：
{ "kind":"body", "runs":[
  { "text":"H" },
  { "text":"2", "subscript":true },
  { "text":"O" }
] }

// ✅ 数学表达式 x² + y² = r²
{ "kind":"equation", "equation":{
  "displayMode":"display",
  "xml":"<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>+</m:t></m:r><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>=</m:t></m:r><m:sSup><m:e><m:r><m:t>r</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>"
} }
```

## 复合样例

完整 RenderSpec 见 [`scripts/sample-equation-render-spec.json`](../scripts/sample-equation-render-spec.json)，覆盖：

1. 二次方程求根公式（分式 + 根号 + 上标嵌套）
2. 高斯积分（带上下限的大型算子）
3. 求和级数（`∑` + 上下限 + 分式 + 上标）
4. 2×2 矩阵
5. 多行推导（`<m:eqArr>` 对齐）
