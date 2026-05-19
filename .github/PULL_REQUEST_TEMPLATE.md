<!-- 感谢 PR！请填写以下信息，便于评审。 -->

## 变更摘要

<!-- 一句话说明这个 PR 做了什么、解决了什么问题。 -->

## 变更类型

- [ ] 🐛 Bug 修复（不破坏现有行为）
- [ ] ✨ 新功能（不破坏现有行为）
- [ ] 💥 破坏性变更（影响 decision JSON 结构、CLI 参数、文件路径或默认行为）
- [ ] 📝 仅文档 / 规则
- [ ] 🔧 构建 / CI / 杂项

## 影响范围

- [ ] `SKILL.md`
- [ ] `references/`
- [ ] `engine/src/`（C# 引擎）
- [ ] `scripts/`（含 validator / 样例 JSON）
- [ ] `agents/`
- [ ] `.github/`（CI / Issue 模板）

## 验证方式

<!-- 详细描述你怎么验证这个 PR 是好的。 -->

- [ ] `dotnet build engine/src/docx-auto-template-engine.csproj -c Release` 通过
- [ ] `python scripts/validate_decision.py scripts/sample-decision.json` 输出 `[OK]`
- [ ] 手工跑了 `render` / `apply`，生成的 `.docx` 用 Word/WPS 打开正常（如适用）
- [ ] 其他：

## 兼容性 / 迁移

<!-- 如果改了 decision schema、CLI、模板默认值，请说明老用户怎么迁移。 -->

## 关联 Issue / 讨论

<!-- e.g. closes #123, refs #456 -->

## 清单

- [ ] 已阅读 [CONTRIBUTING.md](../CONTRIBUTING.md)
- [ ] 已更新 `CHANGELOG.md` 的 `[Unreleased]` 节（如有用户可见变更）
- [ ] 已更新对应 `references/*.md`（如改了 LLM 行为）
- [ ] 已更新 / 新增样例 JSON（如改了 schema）
- [ ] commit message 遵循 Conventional Commits
