# docx-smart-format

> LLM-driven Word document formatter — let the model decide the style; let a local engine write the XML.

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20x64-lightgrey)](#平台与依赖)
[![Skill: Claude Code / Codex](https://img.shields.io/badge/skill-Claude%20Code%20%2F%20Codex-8A2BE2)](SKILL.md)
[![Release](https://img.shields.io/github/v/release/yaya200325/docx-smart-format)](https://github.com/yaya200325/docx-smart-format/releases)

中文为主、英文同步的 Skill。把"用 Word 写论文/报告/纪要时的排版判断"交给 LLM，把"操作 OOXML 的细节"交给本地引擎，避免 LLM 直接拼 XML。

## 特性

- **两种模式**：基于现有 `.docx` 重排（analyze → decide → apply），或从零生成新 `.docx`（render）。
- **LLM 只出决策、不出 XML**：分四层（文档级 / 块级 / 行内 / 对象级）输出 JSON，引擎写回。
- **覆盖排版痛点**：分节、奇偶页眉、STYLEREF 字段、`pgNumType` + `PAGE` 字段、参考文献条目书签 + 上标 REF 交叉引用、横向页竖排页码（`wps:wsp` + VML 双轨）。
- **本地、离线**：skill 不联网，所有模板规则在 `references/`。
- **可校验**：`scripts/validate_decision.py` 校验决策 JSON 结构。

## 平台与依赖

- **当前仅支持 Windows x64**。`engine/runtime/docx-auto-template-engine.exe` 是 .NET self-contained 部署。macOS / Linux 暂未发布预编译产物；可按 `CONTRIBUTING.md` 自行构建对应 RID。
- Skill 自身运行不需要 .NET SDK，只要 runtime 可执行文件存在即可。
- 重新构建引擎需要 .NET SDK 8.0+。

## 安装

### 方式 A：下载 Release（推荐）

1. 到 [Releases](https://github.com/yaya200325/docx-smart-format/releases) 下载最新版本的 `docx-smart-format-<version>-win-x64.zip`。
2. 解压到本机 skill 目录：
   - Claude Code：`~/.claude/skills/docx-smart-format/`
   - Codex：`~/.codex/skills/docx-smart-format/`
   - （Windows 实际路径为 `C:\Users\<USER>\.claude\skills\...` 或 `C:\Users\<USER>\.codex\skills\...`）
3. 解压后目录里应当包含 `engine/runtime/docx-auto-template-engine.exe`。

### 方式 B：克隆仓库 + 本地构建

适用于 Linux/macOS 用户或希望自行编译引擎的开发者：

```powershell
git clone https://github.com/yaya200325/docx-smart-format.git
cd docx-smart-format
# Windows
.\build.ps1
# Linux / macOS
./build.sh
```

`build.ps1` / `build.sh` 会调用 `dotnet publish` 把 `engine/src/` 编译并发布到 `engine/runtime/`。脚本默认 RID 为 `win-x64`，可用 `-Rid linux-x64` / `--rid osx-arm64` 等改写。

> ⚠️ Linux / macOS RID 当前未在 CI 自动验证，引擎对非 Windows 平台的兼容性尚未官方背书；如遇问题欢迎开 Issue。

## 快速开始

### 模式一：重排现有 `.docx`

```powershell
$skill = "$env:USERPROFILE\.claude\skills\docx-smart-format"

# 1. 分析源文档
& "$skill\engine\runtime\docx-auto-template-engine.exe" `
    analyze --input  .\input.docx `
            --output .\analysis.json

# 2. LLM 根据 analysis.json 生成 decision.json（结构见 references/decision-schema.md）
# 3. 写回
& "$skill\engine\runtime\docx-auto-template-engine.exe" `
    apply --source   .\input.docx `
          --decision .\decision.json `
          --output   .\result.docx `
          --template .\template.docx
```

### 模式二：从零生成

```powershell
$skill = "$env:USERPROFILE\.claude\skills\docx-smart-format"

& "$skill\engine\runtime\docx-auto-template-engine.exe" `
    render --spec   .\render-spec.json `
           --output .\result.docx `
           --template-preset builtin-undergraduate-thesis
```

毕业论文场景下 `--template-preset builtin-undergraduate-thesis` 会自动启用 `--normalize-references`。详见 [`references/reference-normalization.md`](references/reference-normalization.md)。

### 校验决策 JSON

```powershell
python scripts\validate_decision.py decision.json
```

通过会打印 `[OK] decision.json 结构合法` 并列出关键计数；错误会以 `[ERROR]` 前缀输出并以非零退出码结束。

## 目录结构

```
docx-smart-format/
├── SKILL.md                       # Skill 主文档（LLM 入口）
├── agents/openai.yaml             # Codex/OpenAI agent 元数据
├── references/                    # 模板规则、能力清单、深度规范
│   ├── decision-schema.md
│   ├── template-catalog.md        # 字体/字号唯一权威
│   ├── local-capability-map.md
│   ├── header-footer-odd-even.md
│   ├── reference-normalization.md
│   └── landscape-vertical-pagenum.md
├── scripts/
│   ├── sample-decision.json
│   ├── sample-render-spec.json
│   ├── sample-chem-render-spec.json
│   └── validate_decision.py
└── engine/
    ├── runtime/                   # 预编译引擎（gitignore，发布到 Release）
    └── src/                       # 引擎源代码（C# / .NET 8）
```

## 文档导航

- LLM 怎么用：从 [`SKILL.md`](SKILL.md) 开始。
- 字段定义：[`references/decision-schema.md`](references/decision-schema.md)。
- 字体字号默认值：[`references/template-catalog.md`](references/template-catalog.md)。
- 奇偶页眉与页码：[`references/header-footer-odd-even.md`](references/header-footer-odd-even.md)。
- 参考文献交叉引用：[`references/reference-normalization.md`](references/reference-normalization.md)。
- 横向页竖排页码：[`references/landscape-vertical-pagenum.md`](references/landscape-vertical-pagenum.md)。
- 引擎可调用能力清单：[`references/local-capability-map.md`](references/local-capability-map.md)。

## 贡献

欢迎 Issue 与 PR：

- 贡献规则 → [CONTRIBUTING.md](CONTRIBUTING.md)
- 行为准则 → [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- 安全问题 → [SECURITY.md](SECURITY.md)（请不要在公开 Issue 中提交安全漏洞）
- 变更历史 → [CHANGELOG.md](CHANGELOG.md)

## 许可证

Apache License 2.0。详见 [`LICENSE`](LICENSE) 与 [`NOTICE`](NOTICE)。
