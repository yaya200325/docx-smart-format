#!/usr/bin/env python3
"""Validate a docx-smart-format decision.json against the current schema.

Schema reference: references/decision-schema.md
"""
import json
import sys
from pathlib import Path
from typing import Any


VALID_DOC_TYPES = {
    "academic-thesis",
    "experiment-report",
    "company-report",
    "meeting-minutes",
    "general-formal-doc",
    "custom",
}

VALID_KINDS = {"paragraph", "table", "image", "equation"}
VALID_VERTICAL_ALIGN = {"baseline", "superscript", "subscript"}
VALID_ROLES = {
    "title",
    "heading1",
    "heading2",
    "heading3",
    "body",
    "bullet",
    "numbered",
    "quote",
    "figureCaption",
    "tableCaption",
    "equation",
    "appendix",
    "reference",
}

VALID_HF_KIND = {"header", "footer"}
VALID_HF_TYPE = {"default", "even", "first"}
VALID_HF_ACTION = {"useTemplate", "clear", "replace"}
VALID_SECTION_TYPE = {"continuous", "nextPage", "oddPage", "evenPage"}
VALID_PGNUM_FORMAT = {
    "decimal",
    "upperRoman",
    "lowerRoman",
    "upperLetter",
    "lowerLetter",
}

REQUIRED_TOP = (
    "docType",
    "templateId",
    "confidence",
    "documentRules",
    "blockDecisions",
    "runDecisions",
    "equationDecisions",
    "assetDecisions",
)


class Reporter:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def error(self, msg: str) -> None:
        self.errors.append(msg)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)


def check_top_level(data: dict[str, Any], r: Reporter) -> None:
    for key in REQUIRED_TOP:
        if key not in data:
            r.error(f"缺少顶层字段: {key}")

    if "globalRules" in data:
        r.warn(
            "顶层出现旧字段 `globalRules`，当前 schema 已改名为 `documentRules`，"
            "请删除残留的 `globalRules` 键。"
        )

    doc_type = data.get("docType")
    if doc_type is not None and doc_type not in VALID_DOC_TYPES:
        r.error(f"非法 docType: {doc_type}")

    template_id = data.get("templateId")
    if not isinstance(template_id, str) or not template_id.strip():
        r.error("templateId 不能为空")

    confidence = data.get("confidence")
    if not isinstance(confidence, (int, float)) or not 0 <= confidence <= 1:
        r.error("confidence 必须是 0 到 1 之间的数字")


def check_document_rules(data: dict[str, Any], r: Reporter) -> bool:
    rules = data.get("documentRules")
    if not isinstance(rules, dict):
        r.error("documentRules 必须是对象")
        return False
    return bool(rules.get("oddEvenDifferent"))


def check_block_decisions(data: dict[str, Any], r: Reporter) -> None:
    items = data.get("blockDecisions")
    if not isinstance(items, list):
        r.error("blockDecisions 必须是数组")
        return

    for index, item in enumerate(items, start=1):
        for key in ("path", "kind", "semanticRole", "targetStyleKey"):
            if key not in item:
                r.error(f"blockDecisions[{index}] 缺少字段: {key}")

        kind = item.get("kind")
        if kind is not None and kind not in VALID_KINDS:
            r.error(f"blockDecisions[{index}] kind 非法: {kind}")

        role = item.get("semanticRole")
        if role is not None and role not in VALID_ROLES:
            r.error(f"blockDecisions[{index}] semanticRole 非法: {role}")


def check_run_decisions(data: dict[str, Any], r: Reporter) -> None:
    items = data.get("runDecisions")
    if not isinstance(items, list):
        r.error("runDecisions 必须是数组")
        return

    for index, item in enumerate(items, start=1):
        if "runPath" not in item:
            r.error(f"runDecisions[{index}] 缺少字段: runPath")

        vertical_align = item.get("verticalAlign")
        if vertical_align is not None and vertical_align not in VALID_VERTICAL_ALIGN:
            r.error(f"runDecisions[{index}] verticalAlign 非法: {vertical_align}")


def check_equation_decisions(data: dict[str, Any], r: Reporter) -> None:
    items = data.get("equationDecisions")
    if not isinstance(items, list):
        r.error("equationDecisions 必须是数组")
        return

    for index, item in enumerate(items, start=1):
        for key in ("path", "kind", "semanticRole", "targetStyleKey"):
            if key not in item:
                r.error(f"equationDecisions[{index}] 缺少字段: {key}")

        if item.get("kind") != "equation":
            r.error(f"equationDecisions[{index}] kind 必须为 equation")

        if item.get("semanticRole") != "equation":
            r.error(f"equationDecisions[{index}] semanticRole 必须为 equation")

        numbering = item.get("numbering")
        if isinstance(numbering, dict) and numbering.get("enabled") is True:
            if not numbering.get("format"):
                r.warn(
                    f"equationDecisions[{index}] numbering.enabled=true 但缺少 format，"
                    "引擎可能使用默认 `(n)`。"
                )


def check_asset_decisions(data: dict[str, Any], r: Reporter) -> None:
    items = data.get("assetDecisions")
    if not isinstance(items, list):
        r.error("assetDecisions 必须是数组")
        return

    for index, item in enumerate(items, start=1):
        for key in ("path", "kind"):
            if key not in item:
                r.error(f"assetDecisions[{index}] 缺少字段: {key}")


def check_header_footer_decisions(
    data: dict[str, Any], r: Reporter
) -> dict[str, dict[str, Any]]:
    """Validate headerFooterDecisions and return a {decisionKey: decision} map."""
    items = data.get("headerFooterDecisions", [])
    if not isinstance(items, list):
        r.error("headerFooterDecisions 必须是数组（可省略，但若提供必须是数组）")
        return {}

    key_map: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(items, start=1):
        prefix = f"headerFooterDecisions[{index}]"

        key = item.get("decisionKey")
        if not isinstance(key, str) or not key.strip():
            r.error(f"{prefix} 缺少非空 decisionKey")
        elif key in key_map:
            r.error(f"{prefix} decisionKey 重复: {key}")
        else:
            key_map[key] = item

        kind = item.get("kind")
        if kind not in VALID_HF_KIND:
            r.error(f"{prefix} kind 非法: {kind!r}，应为 header / footer")

        hf_type = item.get("type")
        if hf_type not in VALID_HF_TYPE:
            if hf_type == "odd":
                r.warn(
                    f"{prefix} type=\"odd\" 不是合法 OOXML 值，"
                    "请用 \"default\" 并在 documentRules 设 oddEvenDifferent=true。"
                )
            else:
                r.error(f"{prefix} type 非法: {hf_type!r}，应为 default / even / first")

        action = item.get("action")
        if action is not None and action not in VALID_HF_ACTION:
            r.error(f"{prefix} action 非法: {action!r}")

        if action == "replace" and not item.get("sourcePath"):
            r.error(f"{prefix} action=replace 必须提供 sourcePath")

    return key_map


def check_section_decisions(
    data: dict[str, Any],
    r: Reporter,
    hf_keys: dict[str, dict[str, Any]],
    odd_even_required: bool,
) -> None:
    items = data.get("sectionDecisions", [])
    if not isinstance(items, list):
        r.error("sectionDecisions 必须是数组（可省略，但若提供必须是数组）")
        return

    referenced_types_header: list[set[str]] = []
    referenced_types_footer: list[set[str]] = []

    for index, item in enumerate(items, start=1):
        prefix = f"sectionDecisions[{index}]"

        if not item.get("afterBlockId") and not item.get("afterBlockPath"):
            r.error(f"{prefix} 必须提供 afterBlockId 或 afterBlockPath")

        st = item.get("sectionType")
        if st is not None and st not in VALID_SECTION_TYPE:
            r.error(f"{prefix} sectionType 非法: {st!r}")

        header_keys = _collect_keys(
            item, "headerDecisionKey", "headerDecisionKeys"
        )
        footer_keys = _collect_keys(
            item, "footerDecisionKey", "footerDecisionKeys"
        )

        header_types = _resolve_types(
            header_keys, hf_keys, expected_kind="header", prefix=f"{prefix}.header",
            reporter=r,
        )
        footer_types = _resolve_types(
            footer_keys, hf_keys, expected_kind="footer", prefix=f"{prefix}.footer",
            reporter=r,
        )

        referenced_types_header.append(header_types)
        referenced_types_footer.append(footer_types)

        pg = item.get("pgNumType")
        if pg is not None:
            if not isinstance(pg, dict):
                r.error(f"{prefix} pgNumType 必须是对象")
            else:
                fmt = pg.get("format")
                if fmt is not None and fmt not in VALID_PGNUM_FORMAT:
                    r.error(f"{prefix} pgNumType.format 非法: {fmt!r}")
                start = pg.get("start")
                if start is not None and not isinstance(start, int):
                    r.error(f"{prefix} pgNumType.start 必须是整数")

    if odd_even_required:
        for index, types in enumerate(referenced_types_header, start=1):
            if types and not ({"default", "even"} <= types):
                r.error(
                    f"documentRules.oddEvenDifferent=true，但 sectionDecisions[{index}] "
                    f"的 header 仅绑定 {sorted(types)}，必须同时绑定 default + even。"
                )


def _collect_keys(item: dict[str, Any], single: str, plural: str) -> list[str]:
    keys: list[str] = []
    s = item.get(single)
    if isinstance(s, str) and s:
        keys.append(s)
    p = item.get(plural)
    if isinstance(p, list):
        keys.extend(k for k in p if isinstance(k, str) and k)
    return keys


def _resolve_types(
    keys: list[str],
    hf_keys: dict[str, dict[str, Any]],
    expected_kind: str,
    prefix: str,
    reporter: Reporter,
) -> set[str]:
    types: set[str] = set()
    for key in keys:
        decision = hf_keys.get(key)
        if decision is None:
            reporter.error(
                f"{prefix} 引用了未注册的 decisionKey={key!r}，"
                "请在 headerFooterDecisions 中声明。"
            )
            continue
        if decision.get("kind") != expected_kind:
            reporter.error(
                f"{prefix} 引用了 kind={decision.get('kind')!r} 的部件，"
                f"应为 {expected_kind!r}。"
            )
            continue
        t = decision.get("type")
        if isinstance(t, str):
            types.add(t)
    return types


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: validate_decision.py <decision.json>")
        return 1

    path = Path(sys.argv[1])
    if not path.is_file():
        print(f"[ERROR] 文件不存在: {path}")
        return 1

    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        print(f"[ERROR] JSON 解析失败: {exc}")
        return 1

    r = Reporter()

    check_top_level(data, r)
    odd_even_required = check_document_rules(data, r)
    check_block_decisions(data, r)
    check_run_decisions(data, r)
    check_equation_decisions(data, r)
    check_asset_decisions(data, r)
    hf_keys = check_header_footer_decisions(data, r)
    check_section_decisions(data, r, hf_keys, odd_even_required)

    for w in r.warnings:
        print(f"[WARN] {w}")
    for e in r.errors:
        print(f"[ERROR] {e}")

    if r.errors:
        return 1

    print("[OK] decision.json 结构合法")
    print(f"docType={data.get('docType')}")
    print(f"templateId={data.get('templateId')}")
    print(f"blockDecisions={len(data.get('blockDecisions', []))}")
    print(f"runDecisions={len(data.get('runDecisions', []))}")
    print(f"equationDecisions={len(data.get('equationDecisions', []))}")
    print(f"assetDecisions={len(data.get('assetDecisions', []))}")
    print(f"headerFooterDecisions={len(data.get('headerFooterDecisions', []))}")
    print(f"sectionDecisions={len(data.get('sectionDecisions', []))}")
    if r.warnings:
        print(f"warnings={len(r.warnings)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
