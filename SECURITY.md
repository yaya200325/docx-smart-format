# 安全政策 / Security Policy

## 受支持的版本

本项目目前处于早期阶段，仅对主分支 (`main`) 与最新发布的小版本提供安全更新。

| 版本     | 是否提供安全更新 |
|----------|------------------|
| 0.1.x    | ✅               |
| < 0.1.0  | ❌（请升级）     |

## 报告安全漏洞

**请不要通过公开 Issue 报告安全漏洞。** 公开漏洞前请先与维护者私下沟通，给予合理时间修复。

报告方式（按优先级）：

1. 通过 GitHub 的 [Private vulnerability reporting](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability) 功能提交：进入仓库 → Security tab → Report a vulnerability。
2. 若上述渠道不可用，可在 GitHub 上私聊 [@yaya200325](https://github.com/yaya200325)。

请在报告中尽量包含：

- 漏洞类型、受影响的组件（如 `engine/runtime`、`scripts/validate_decision.py`、某个 `references/*.md` 中的规则）。
- 复现步骤、最小复现样本（若涉及恶意 `.docx`，请尽可能脱敏后再附）。
- 影响范围（信息泄露 / 任意文件写入 / 拒绝服务 / 其他）。
- 建议的缓解方案（可选）。

## 响应时间预期

| 阶段           | 目标响应时间           |
|----------------|------------------------|
| 首次确认收到   | 5 个工作日以内         |
| 初步评估       | 14 个工作日以内        |
| 修复或缓解方案 | 视严重程度，最长 90 天 |

修复发布后会在 [CHANGELOG.md](CHANGELOG.md) 与 GitHub Security Advisory 公开披露。

## 安全相关的设计约束

理解以下边界有助于判断是否构成漏洞：

- **Skill 不联网**：所有规则在本地 `references/` 与 `engine/runtime/`，不发起出站请求。如发现 skill 在运行期间联网访问外部服务，请按漏洞报告处理。
- **引擎读取 `.docx`**：本质是解压 ZIP + 解析 XML。恶意构造的 `.docx`（zip-slip、XML 外部实体、巨型膨胀文件等）若导致引擎写入受控目录之外、读取任意文件或长时间挂起，属于漏洞。
- **`scripts/validate_decision.py`**：仅做 JSON 结构校验，不执行 decision 内容。如果发现存在通过决策 JSON 触发任意命令执行的路径，属于漏洞。
- **decision JSON 中 `sourcePath`**：用于引用页眉/页脚的 XML 片段。引擎应限制其只能解析为决策文件同目录及子目录下的资源；若可逃逸到任意路径，请报告。

## 致谢

感谢所有负责任披露漏洞的研究者。修复发布时会在 advisory 中署名（如本人同意）。
