# -*- coding: utf-8 -*-
"""ra-doc-assist 与 SKILL.md 对齐的控制台提示与轻量校验（不依赖 LLM）。"""
from __future__ import annotations

import json
import re
import sys
from typing import Any, List, Optional, Tuple

# 与 SKILL.md「执行铁律」同步；修改 SKILL 时请同步更新本常量
SKILL_IRON_RULES = (
    "【ra-doc-assist 执行铁律】",
    "1) 用 Read 打开本仓库 REFERENCE.md：「LLM 整合规则（执行版）」「refined.json 输出结构」「输出内容结构与整段示例」",
    "2) SOP 整合固定三步，禁止跳步：integrate_sop_method.py --extract-only → python refine_extracted.py → integrate_sop_method.py --from-json",
    "3) 禁止另写脚本生成/篡改 refined.json；中间 JSON 每次任务重新生成，默认放 output/",
    "4) principle（原理）须与 SOP「实验原理」原文一致（refine_extracted 已保留全文；可用 --strict-principle 强校验）",
    "5) Word 版式遵守 SKILL.md 文末「CTD 资料模板格式要求」§1～§12；写回以公司标准模版样式为准",
)


def print_skill_iron_rules(stream=sys.stdout) -> None:
    for line in SKILL_IRON_RULES:
        print(line, file=stream)
    print(file=stream)


def _reconstruct_blob(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, str):
        return val
    if isinstance(val, list):
        if len(val) == 0:
            return ""
        if isinstance(val[0], str) and len(val[0]) == 1:
            return "".join(val)
        return "\n".join(str(x) for x in val)
    return str(val)


def _strip_principle_prefix(s: str) -> str:
    return re.sub(r"^原理[：:]\s*", "", (s or "").strip())


def _collapse_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def verify_principle_verbatim(
    sop_raw_principle: Any, refined_principle: Any
) -> Tuple[bool, str]:
    """
    检查 refined.principle 是否与 SOP 原理正文一致（空白归一后比较）。
    refined_principle: list[str] 或 str
    """
    sop = _strip_principle_prefix(_reconstruct_blob(sop_raw_principle))
    if not sop.strip():
        return True, ""

    if refined_principle is None:
        return False, "SOP 含实验原理，但 refined.principle 缺失"

    if isinstance(refined_principle, list):
        ref = "\n".join(str(x) for x in refined_principle if str(x).strip())
    else:
        ref = str(refined_principle)
    ref = _strip_principle_prefix(ref)

    if not ref.strip():
        return False, "SOP 含实验原理，但 refined.principle 为空"

    a, b = _collapse_ws(sop), _collapse_ws(ref)
    if a == b:
        return True, ""
    return (
        False,
        "原理正文与 SOP 归一化后不逐字一致（请检查是否手工改写了 refined.principle）",
    )


def validate_refined_principles_in_files(
    refined_path: str, extracted_path: str
) -> Tuple[bool, List[str]]:
    """对 refined.json 中每个 method 校验 principle；需可读的 extracted.json。"""
    errors: List[str] = []
    with open(refined_path, "r", encoding="utf-8") as f:
        refined = json.load(f)
    with open(extracted_path, "r", encoding="utf-8") as f:
        ext = json.load(f)

    r_methods = refined.get("methods") or []
    e_methods = ext.get("methods") or []
    if len(r_methods) != len(e_methods):
        errors.append(
            f"methods 数量不一致: refined={len(r_methods)} extracted={len(e_methods)}"
        )

    for i, (rm, em) in enumerate(zip(r_methods, e_methods)):
        name = rm.get("name") or em.get("name") or f"#{i + 1}"
        sop_raw = em.get("sop_raw") or {}
        refined_sec = (rm.get("refined") or {}).get("principle")
        ok, msg = verify_principle_verbatim(
            sop_raw.get("principle"), refined_sec
        )
        if not ok:
            errors.append(f"{name}: {msg}")

    return len(errors) == 0, errors


def main_validate(argv: Optional[List[str]] = None) -> int:
    """CLI: python -m ra_compliance refined.json extracted.json"""
    args = argv if argv is not None else sys.argv[1:]
    if len(args) < 2:
        print(
            "用法: python ra_compliance.py <refined.json> <extracted.json>",
            file=sys.stderr,
        )
        return 2
    ok, errs = validate_refined_principles_in_files(args[0], args[1])
    if ok:
        print("[OK] principle 与 SOP 逐字校验通过")
        return 0
    print("[FAIL] principle 校验:", file=sys.stderr)
    for e in errs:
        print(f"  - {e}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main_validate())
